import os
import zipfile
import pandas as pd
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, CallbackQueryHandler, ContextTypes, filters

USER_SETTINGS = {}

# Целевой формат колонок как в файле Bookings_unique_part_7.csv
TARGET_COLUMNS = [
    'id', 'fio', 'check_in', 'check_out', 'price', 'currency', 
    'email', 'phone', 'hotel_name', 'address', 'image', 'urls'
]

# Словари синонимов, чтобы бот сам понимал "другие форматы" 
COLUMN_ALIASES = {
    'id': ['id', 'booking_id', 'reservation', 'номер', 'reservation_id'],
    'fio': ['fio', 'name', 'guest', 'имя', 'фио', 'клиент', 'guest_name', 'full_name'],
    'check_in': ['check_in', 'checkin', 'заезд', 'arrival', 'date_in'],
    'check_out': ['check_out', 'checkout', 'выезд', 'departure', 'date_out'],
    'price': ['price', 'amount', 'цена', 'сумма', 'total'],
    'currency': ['currency', 'валюта', 'curr'],
    'email': ['email', 'почта', 'e-mail'],
    'phone': ['phone', 'телефон', 'tel', 'mobile'],
    'hotel_name': ['hotel_name', 'hotel', 'отель', 'гостиница'],
    'address': ['address', 'адрес', 'location'],
    'image': ['image', 'фото', 'picture'],
    'urls': ['urls', 'url', 'ссылка', 'link']
}

def standardize_dataframe(df):
    """Подгоняет исходный DataFrame под единый стандартный вид."""
    new_df = pd.DataFrame()
    
    # Переводим названия исходных колонок в нижний регистр для сравнения
    df_cols_lower = {str(col).strip().lower(): col for col in df.columns}
    
    for target in TARGET_COLUMNS:
        matched = False
        # Ищем совпадения по синонимам
        for alias in COLUMN_ALIASES.get(target, [target]):
            if alias in df_cols_lower:
                original_col_name = df_cols_lower[alias]
                new_df[target] = df[original_col_name]
                matched = True
                break
        
        # Если колонка не нашлась в исходнике, создаем её пустой
        if not matched:
            new_df[target] = pd.NA

    # --- Подстановка стандартных текстов из Bookings_unique_part_7.csv ---
    
    # 1. Генерируем hotel_name: "Hotel confirmation for reservation {id}" если он пуст
    mask_hotel = new_df['hotel_name'].isna() & new_df['id'].notna() & (new_df['id'] != '')
    clean_ids = new_df['id'].fillna('').astype(str).str.replace(r'\.0$', '', regex=True)
    new_df.loc[mask_hotel, 'hotel_name'] = "Hotel confirmation for reservation " + clean_ids[mask_hotel]

    # 2. Стандартный текст address
    mask_addr = new_df['address'].isna()
    new_df.loc[mask_addr, 'address'] = "You need to confirm your booking. This is required for verification purposes."

    # 3. Стандартная картинка image
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
        "Выберите количество строк в одном .xlsx файле.\n"
        "Бот автоматически подгонит колонки под нужный формат и пропустит битые строки.", 
        reply_markup=reply_markup
    )

async def set_rows(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    chunk_size = int(query.data.split('_')[1])
    USER_SETTINGS[query.from_user.id] = chunk_size
    await query.edit_message_text(f"Лимит установлен: {chunk_size} строк. Присылайте CSV, TXT или ZIP файл.")

async def process_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    chunk_size = USER_SETTINGS.get(user_id, 1000)
    
    doc = update.message.document
    file_name = doc.file_name.lower()
    file = await context.bot.get_file(doc.file_id)
    file_path = f"temp_{doc.file_id}_{doc.file_name}"
    await file.download_to_drive(file_path)
    
    msg = await update.message.reply_text("Обрабатываю данные и подгоняю формат колонок...")

    try:
        dfs = []
        if file_name.endswith('.zip'):
            with zipfile.ZipFile(file_path, 'r') as zip_ref:
                for name in zip_ref.namelist():
                    if name.lower().endswith(('.csv', '.txt')) and not name.startswith('__MACOSX'):
                        with zip_ref.open(name) as f:
                            # on_bad_lines='skip' игнорирует битые строки!
                            df = pd.read_csv(f, sep=None, engine='python', dtype=str, on_bad_lines='skip')
                            dfs.append(df)
        elif file_name.endswith(('.csv', '.txt')):
            df = pd.read_csv(file_path, sep=None, engine='python', dtype=str, on_bad_lines='skip')
            dfs.append(df)

        if not dfs:
            await msg.edit_text("Подходящие CSV/TXT файлы не найдены.")
            return

        # Склеиваем всё в один DataFrame
        combined_df = pd.concat(dfs, ignore_index=True)
        
        # Запускаем нашу умную подгонку
        standardized_df = standardize_dataframe(combined_df)

        total_rows = len(standardized_df)
        parts_count = (total_rows + chunk_size - 1) // chunk_size
        
        for part in range(parts_count):
            start_idx = part * chunk_size
            end_idx = start_idx + chunk_size
            chunk = standardized_df.iloc[start_idx:end_idx]
            
            out_name = f"output_part_{part + 1}.xlsx"
            # Сохраняем в XLSX без технического индекса
            chunk.to_excel(out_name, index=False)
            
            with open(out_name, 'rb') as f:
                await update.message.reply_document(document=f, filename=out_name)
            os.remove(out_name)

        await msg.edit_text(f"Готово! Обработано строк: {total_rows}. Выгружено строго по шаблону.")

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
