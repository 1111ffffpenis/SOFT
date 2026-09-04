import os
import zipfile
import re
import pandas as pd
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, CallbackQueryHandler, ContextTypes, filters
from openpyxl.styles import Font, PatternFill

USER_SETTINGS = {}

# Единый целевой формат колонок
TARGET_COLUMNS = [
    'id', 'fio', 'check_in', 'check_out', 'price', 'currency', 
    'email', 'phone', 'hotel_name', 'address', 'image', 'urls'
]

# Фиксированные гарантированные ширины колонок для защиты от наложения текста
COLUMN_WIDTHS = {
    'A': 18, # id
    'B': 28, # fio
    'C': 14, # check_in
    'D': 14, # check_out
    'E': 14, # price
    'F': 14, # currency
    'G': 32, # email
    'H': 18, # phone
    'I': 35, # hotel_name
    'J': 45, # address
    'K': 35, # image
    'L': 25  # urls
}

COLUMN_ALIASES = {
    'id': ['id', 'resv_id', 'booking_id', 'reservation_id', 'reservation', 'order_id', 'номер_брони', 'номер', 'src_id', 'room_id'],
    'fio': ['fio', 'guest_name', 'customer_name', 'guest', 'name', 'full_name', 'фио', 'имя', 'клиент'],
    'check_in': ['check_in', 'check_in_date', 'checkin_date', 'checkin', 'arrival_date', 'arrival', 'заезд', 'дата_заезда', 'date_in', 'start_date'],
    'check_out': ['check_out', 'check_out_date', 'checkout_date', 'checkout', 'departure_date', 'departure', 'выезд', 'дата_выезда', 'date_out', 'end_date'],
    'price': ['price', 'total_price', 'total_amount', 'price_total', 'amount', 'total', 'cost', 'цена', 'сумма', 'стоимость'],
    'currency': ['currency', 'curr', 'valuta', 'валюта'],
    'email': ['email', 'mail', 'e-mail', 'почта'],
    'phone': ['phone', 'telephone', 'mobile', 'tel', 'телефон', 'номер_телефона'],
    'hotel_name': ['hotel_name', 'hotel', 'property_name', 'property', 'отель', 'гостиница'],
    'address': ['address', 'location', 'адрес'],
    'image': ['image', 'img', 'photo', 'picture', 'фото'],
    'urls': ['urls', 'url', 'link', 'links', 'ссылка', 'ссылки']
}

def robust_read_csv(file_path):
    """Чтение CSV/TXT с подбором разделителя для предотвращения сжатия столбцов."""
    try:
        df = pd.read_csv(file_path, sep=None, engine='python', dtype=str, on_bad_lines='skip')
        if len(df.columns) > 1:
            return df
    except Exception:
        pass
    
    for sep in [';', ',', '\t', '|']:
        try:
            df = pd.read_csv(file_path, sep=sep, dtype=str, on_bad_lines='skip')
            if len(df.columns) > 1:
                return df
        except Exception:
            continue
            
    return pd.read_csv(file_path, sep=',', dtype=str, on_bad_lines='skip')

def standardize_dataframe(df):
    """Приводит колонки к единому стандарту и заполняет недостающие поля."""
    new_df = pd.DataFrame()
    
    df_cols = {}
    for col in df.columns:
        norm_col = re.sub(r'[\s\.\-]+', '_', str(col).strip().lower())
        df_cols[norm_col] = col
        
    used_cols = set()
    for target in TARGET_COLUMNS:
        matched_col = None
        aliases = COLUMN_ALIASES.get(target, [target])
        
        for alias in aliases:
            if alias in df_cols and df_cols[alias] not in used_cols:
                matched_col = df_cols[alias]
                break
                
        if not matched_col:
            for alias in aliases:
                for norm_col, orig_col in df_cols.items():
                    if orig_col in used_cols:
                        continue
                    if re.search(r'(^|_)' + re.escape(alias) + r'($|_)', norm_col):
                        matched_col = orig_col
                        break
                if matched_col:
                    break

        if matched_col:
            new_df[target] = df[matched_col]
            used_cols.add(matched_col)
        else:
            new_df[target] = pd.NA

    # Шаблоны по умолчанию
    mask_hotel = new_df['hotel_name'].isna() & new_df['id'].notna() & (new_df['id'] != '')
    clean_ids = new_df['id'].fillna('').astype(str).str.replace(r'\.0$', '', regex=True)
    new_df.loc[mask_hotel, 'hotel_name'] = "Hotel confirmation for reservation " + clean_ids[mask_hotel]

    mask_addr = new_df['address'].isna()
    new_df.loc[mask_addr, 'address'] = "You need to confirm your booking. This is required for verification purposes."

    mask_img = new_df['image'].isna()
    new_df.loc[mask_img, 'image'] = "https://i.ibb.co/C5dHd4fv/image.png"

    return new_df

def save_excel_perfect(df, filename):
    """Сохраняет XLSX с фиксированными широкими колонками и форматированием."""
    writer = pd.ExcelWriter(filename, engine='openpyxl')
    df.to_excel(writer, index=False, sheet_name='Data')
    worksheet = writer.sheets['Data']
    
    # Задаем гарантированную ширину столбцов
    for col_letter, width in COLUMN_WIDTHS.items():
        worksheet.column_dimensions[col_letter].width = width
        
    # Выделяем шапку серым фоном и жирным шрифтом
    header_fill = PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid")
    header_font = Font(bold=True)
    for cell in worksheet[1]:
        cell.fill = header_fill
        cell.font = header_font

    writer.close()

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
    
    msg = await update.message.reply_text("Обрабатываю файлы и форматирую таблицу...")

    try:
        dfs = []
        if file_name.endswith('.zip'):
            with zipfile.ZipFile(file_path, 'r') as zip_ref:
                for name in zip_ref.namelist():
                    if name.lower().endswith(('.csv', '.txt')) and not name.startswith('__MACOSX'):
                        zip_ref.extract(name, path=".")
                        df = robust_read_csv(name)
                        dfs.append(df)
                        os.remove(name)
        elif file_name.endswith(('.csv', '.txt')):
            df = robust_read_csv(file_path)
            dfs.append(df)

        if not dfs:
            await msg.edit_text("Файлы CSV/TXT не найдены.")
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
            
            out_name = f"output_part_{part + 1}_{current_chunk_rows}rows.xlsx"
            save_excel_perfect(chunk, out_name)
            
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
