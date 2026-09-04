import os
import zipfile
import pandas as pd
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, CallbackQueryHandler, ContextTypes, filters

USER_SETTINGS = {}

# Единый целевой формат колонок
TARGET_COLUMNS = [
    'id', 'fio', 'check_in', 'check_out', 'price', 'currency', 
    'email', 'phone', 'hotel_name', 'address', 'image', 'urls'
]

# Варианты названий колонок в различных выгрузках (включая Euro.csv)
COLUMN_ALIASES = {
    'id': ['resv_id', 'booking_id', 'reservation_id', 'order_id', 'id', 'номер_брони', 'номер', 'src_id', 'room_id'],
    'fio': ['guest_name', 'customer_name', 'fio', 'guest', 'name', 'full_name', 'фио', 'имя', 'клиент'],
    'check_in': ['check_in_date', 'checkin_date', 'check_in', 'checkin', 'arrival_date', 'arrival', 'заезд', 'дата_заезда', 'date_in', 'start_date'],
    'check_out': ['check_out_date', 'checkout_date', 'check_out', 'checkout', 'departure_date', 'departure', 'выезд', 'дата_выезда', 'date_out', 'end_date'],
    'price': ['total_price', 'total_amount', 'price_total', 'price', 'amount', 'total', 'cost', 'цена', 'сумма', 'стоимость'],
    'currency': ['currency', 'curr', 'valuta', 'валюта'],
    'email': ['email', 'mail', 'e-mail', 'почта'],
    'phone': ['phone', 'telephone', 'mobile', 'tel', 'телефон', 'номер_телефона'],
    'hotel_name': ['hotel_name', 'hotel', 'property_name', 'property', 'отель', 'гостиница'],
    'address': ['address', 'location', 'адрес'],
    'image': ['image', 'img', 'photo', 'picture', 'фото'],
    'urls': ['urls', 'url', 'link', 'links', 'ссылка', 'ссылки']
}

def standardize_dataframe(df):
    """Приводит DataFrame к единой структуре и восстанавливает отсутствующие данные."""
    new_df = pd.DataFrame()
    df_cols = {str(col).strip().lower(): col for col in df.columns}
    used_cols = set()
    
    for target in TARGET_COLUMNS:
        matched_col = None
        aliases = COLUMN_ALIASES.get(target, [target])
        
        # 1. Точное совпадение по алиасам
        for alias in aliases:
            if alias in df_cols and df_cols[alias] not in used_cols:
                matched_col = df_cols[alias]
                break
                
        # 2. Поиск по частичному совпадению
        if not matched_col:
            for alias in aliases:
                for col_lower, orig_col in df_cols.items():
                    if orig_col in used_cols:
                        continue
                    if alias in col_lower:
                        matched_col = orig_col
                        break
                if matched_col:
                    break
                    
        if matched_col:
            new_df[target] = df[matched_col]
            used_cols.add(matched_col)
        else:
            new_df[target] = pd.NA

    # Подстановка стандартных текстовых заполнителей
    mask_hotel = new_df['hotel_name'].isna() & new_df['id'].notna() & (new_df['id'] != '')
    clean_ids = new_df['id'].fillna('').astype(str).str.replace(r'\.0$', '', regex=True)
    new_df.loc[mask_hotel, 'hotel_name'] = "Hotel confirmation for reservation " + clean_ids[mask_hotel]

    mask_addr = new_df['address'].isna()
    new_df.loc[mask_addr, 'address'] = "You need to confirm your booking. This is required for verification purposes."

    mask_img = new_df['image'].isna()
    new_df.loc[mask_img, 'image'] = "https://i.ibb.co/C5dHd4fv/image.png"

    return new_df


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    keyboard = [
        [InlineKeyboardButton("1000", callback_data="rows_1000"), InlineKeyboardButton("2000", callback_data="rows_2000")],
        [InlineKeyboardButton("3000", callback_data="rows_3000"), InlineKeyboardButton("5000", callback_data="rows_5000")]
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)
    await update.message.reply_text(
        "Выберите желаемый лимит строк на один .xlsx файл:", 
        reply_markup=reply_markup
    )

async def set_rows(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    chunk_size = int(query.data.split('_')[1])
    USER_SETTINGS[query.from_user.id] = chunk_size
    await query.edit_message_text(f"Лимит строк установлен: {chunk_size}. Присылайте CSV, TXT или ZIP файл.")

async def process_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    chunk_size = USER_SETTINGS.get(user_id, 1000)
    
    doc = update.message.document
    file_name = doc.file_name.lower()
    file = await context.bot.get_file(doc.file_id)
    file_path = f"temp_{doc.file_id}_{doc.file_name}"
    await file.download_to_drive(file_path)
    
    msg = await update.message.reply_text("Обрабатываю файлы и настраиваю структуру...")

    try:
        dfs = []
        if file_name.endswith('.zip'):
            with zipfile.ZipFile(file_path, 'r') as zip_ref:
                for name in zip_ref.namelist():
                    if name.lower().endswith(('.csv', '.txt')) and not name.startswith('__MACOSX'):
                        with zip_ref.open(name) as f:
                            df = pd.read_csv(f, sep=None, engine='python', dtype=str, on_bad_lines='skip')
                            dfs.append(df)
        elif file_name.endswith(('.csv', '.txt')):
            df = pd.read_csv(file_path, sep=None, engine='python', dtype=str, on_bad_lines='skip')
            dfs.append(df)

        if not dfs:
            await msg.edit_text("Файлы CSV/TXT не найдены в сообщении.")
            return

        combined_df = pd.concat(dfs, ignore_index=True)
        standardized_df = standardize_dataframe(combined_df)

        total_rows = len(standardized_df)
        parts_count = (total_rows + chunk_size - 1) // chunk_size
        
        for part in range(parts_count):
            start_idx = part * chunk_size
            end_idx = start_idx + chunk_size
            chunk = standardized_df.iloc[start_idx:end_idx]
            current_chunk_rows = len(chunk)
            
            # Название файла с указанием количества строк
            out_name = f"output_part_{part + 1}_{current_chunk_rows}rows.xlsx"
            chunk.to_excel(out_name, index=False)
            
            with open(out_name, 'rb') as f:
                await update.message.reply_document(
                    document=f, 
                    filename=out_name,
                    caption=f"📄 Файл {part + 1} из {parts_count}\nКоличество строк: {current_chunk_rows}"
                )
            os.remove(out_name)

        await msg.edit_text(f"Готово!\nВсего обработано строк: {total_rows}\nСформировано файлов: {parts_count}")

    except Exception as e:
        await msg.edit_text(f"Ошибка при обработке: {str(e)}")
    finally:
        if os.path.exists(file_path):
            os.remove(file_path)

if __name__ == '__main__':
    bot_token = os.getenv("BOT_TOKEN")
    if not bot_token:
        raise ValueError("BOT_TOKEN variable is not set!")
        
    app = ApplicationBuilder().token(bot_token).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CallbackQueryHandler(set_rows, pattern="^rows_"))
    app.add_handler(MessageHandler(filters.Document.ALL, process_file))
    
    app.run_polling()
