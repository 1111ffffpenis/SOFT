import os
import zipfile
import re
import csv
import pandas as pd
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, CallbackQueryHandler, ContextTypes, filters
from openpyxl.styles import Font, PatternFill

USER_SETTINGS = {}

TARGET_COLUMNS = [
    'id', 'fio', 'check_in', 'check_out', 'price', 'currency', 
    'email', 'phone', 'hotel_name', 'address', 'image', 'urls'
]

COLUMN_WIDTHS = {
    'A': 18, 'B': 28, 'C': 14, 'D': 14, 'E': 14, 'F': 14, 
    'G': 32, 'H': 18, 'I': 35, 'J': 45, 'K': 35, 'L': 25
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

# --- ГЛОБАЛЬНЫЙ СЛОВАРЬ ВСЕХ ВАЛЮТ МИРА ---
CURRENCY_MAP = {
    # Европа
    'euro': 'EUR', 'eur': 'EUR', '€': 'EUR', 'euros': 'EUR',
    'rub': 'RUB', 'ruble': 'RUB', 'руб': 'RUB', 'рубль': 'RUB',
    'gbp': 'GBP', 'pound': 'GBP', '£': 'GBP', 'pounds': 'GBP',
    'chf': 'CHF', 'swiss franc': 'CHF',
    'pln': 'PLN', 'zloty': 'PLN', 'polish zloty': 'PLN',
    'czk': 'CZK', 'koruna': 'CZK', 'czech koruna': 'CZK',
    'huf': 'HUF', 'forint': 'HUF',
    'ron': 'RON', 'leu': 'RON',
    'bgn': 'BGN', 'lev': 'BGN',
    'sek': 'SEK', 'swedish krona': 'SEK',
    'nok': 'NOK', 'norwegian krone': 'NOK',
    'dkk': 'DKK', 'danish krone': 'DKK',
    'rsd': 'RSD', 'serbian dinar': 'RSD',

    # Америка
    'usd': 'USD', 'dollar': 'USD', '$': 'USD', 'us dollar': 'USD', 'dollars': 'USD',
    'cad': 'CAD', 'canadian dollar': 'CAD',
    'mxn': 'MXN', 'mexican peso': 'MXN',
    'brl': 'BRL', 'real': 'BRL', 'brazilian real': 'BRL',
    'ars': 'ARS', 'argentine peso': 'ARS',
    'cop': 'COP', 'colombian peso': 'COP', 'columbian': 'COP', 'columbian peso': 'COP',
    'clp': 'CLP', 'chilean peso': 'CLP',
    'pen': 'PEN', 'sol': 'PEN',
    'crc': 'CRC', 'colón': 'CRC',

    # Азия
    'cny': 'CNY', 'yuan': 'CNY', 'renminbi': 'CNY', 'rmb': 'CNY',
    'jpy': 'JPY', 'yen': 'JPY', '¥': 'JPY',
    'krw': 'KRW', 'won': 'KRW', 'korean won': 'KRW',
    'inr': 'INR', 'indian rupee': 'INR', 'rupees': 'INR', 'indrian rupies': 'INR', 'rupee': 'INR',
    'idr': 'IDR', 'rupiah': 'IDR', 'indonesian rupiah': 'IDR',
    'myr': 'MYR', 'ringgit': 'MYR', 'malaysian ringgit': 'MYR',
    'sgd': 'SGD', 'singapore dollar': 'SGD',
    'php': 'PHP', 'peso': 'PHP', 'philippine peso': 'PHP',
    'thb': 'THB', 'baht': 'THB', 'thai baht': 'THB',
    'vnd': 'VND', 'dong': 'VND', 'vietnamese dong': 'VND',
    'twd': 'TWD', 'new taiwan dollar': 'TWD',
    'hkd': 'HKD', 'hong kong dollar': 'HKD',
    'pkr': 'PKR', 'pakistani rupee': 'PKR',
    'bdt': 'BDT', 'taka': 'BDT',

    # Ближний Восток
    'aed': 'AED', 'dirham': 'AED', 'uae dirham': 'AED',
    'kwd': 'KWD', 'kuwaiti dinar': 'KWD', 'kuwaiti': 'KWD', 'dinar': 'KWD',
    'sar': 'SAR', 'riyal': 'SAR', 'saudi riyal': 'SAR',
    'qar': 'QAR', 'qatari riyal': 'QAR',
    'omr': 'OMR', 'omani rial': 'OMR',
    'bhd': 'BHD', 'bahraini dinar': 'BHD',
    'ils': 'ILS', 'shekel': 'ILS', 'israeli new shekel': 'ILS',
    'jod': 'JOD', 'jordanian dinar': 'JOD',
    'lbp': 'LBP', 'lebanese pound': 'LBP',
    'try': 'TRY', 'lira': 'TRY', 'turkish lira': 'TRY',

    # Африка и Океания
    'aud': 'AUD', 'australian dollar': 'AUD',
    'nzd': 'NZD', 'new zealand dollar': 'NZD',
    'zar': 'ZAR', 'rand': 'ZAR', 'south african rand': 'ZAR',
    'egp': 'EGP', 'egyptian pound': 'EGP',
    'mad': 'MAD', 'moroccan dirham': 'MAD',
    'ngn': 'NGN', 'naira': 'NGN',
    'kes': 'KES', 'kenyan shilling': 'KES',

    # СНГ
    'kzt': 'KZT', 'tenge': 'KZT', 'тенге': 'KZT',
    'byn': 'BYN', 'belarusian ruble': 'BYN',
    'uah': 'UAH', 'hryvnia': 'UAH', 'гривна': 'UAH',
    'uzs': 'UZS', 'som': 'UZS', 'сум': 'UZS',
    'gel': 'GEL', 'lari': 'GEL', 'лари': 'GEL',
    'amd': 'AMD', 'dram': 'AMD', 'драм': 'AMD',
    'azn': 'AZN', 'manat': 'AZN',

    # Крипта
    'btc': 'BTC', 'bitcoin': 'BTC',
    'eth': 'ETH', 'ethereum': 'ETH',
    'usdt': 'USDT', 'tether': 'USDT'
}

def normalize_currency(val):
    if pd.isna(val) or not isinstance(val, str):
        return val
    
    val_lower = val.strip().lower()
    
    # 1. Прямое совпадение
    if val_lower in CURRENCY_MAP:
        return CURRENCY_MAP[val_lower]
        
    # 2. Поиск по подстроке (например, "kuwaiti dinar 200")
    for key, code in CURRENCY_MAP.items():
        if key in val_lower:
            return code
            
    # 3. Если 3 буквы — делаем капсом
    if len(val_lower) == 3:
        return val_lower.upper()
        
    return val.title()


def robust_read_csv(file_path):
    """
    Бронебойное чтение с авто-починкой СМЕЩЕННЫХ строк.
    """
    with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
        first_line = f.readline()
        
    sep = ','
    if ';' in first_line: sep = ';'
    elif '\t' in first_line: sep = '\t'
    elif '|' in first_line: sep = '|'
    
    with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
        reader = csv.reader(f, delimiter=sep)
        rows = list(reader)
        
    if not rows:
        return pd.DataFrame()
        
    header = [str(x).strip() for x in rows[0]]
    expected_len = len(header)
    
    fixed_rows = []
    for row in rows[1:]:
        if not any(row): continue
        
        # ЕСЛИ СТОЛБЦЫ СМЕСТИЛИСЬ ИЗ-ЗА ЛИШНЕЙ ЗАПЯТОЙ
        if len(row) > expected_len:
            diff = len(row) - expected_len
            # Склеиваем лишние данные в одной ячейке (обычно во 2-й колонке FIO), 
            # чтобы остальные важные данные встали на свои законные места.
            left_part = row[:1]
            middle_part = [" ".join(row[1:1+diff+1])]
            right_part = row[1+diff+1:]
            fixed_row = left_part + middle_part + right_part
            fixed_rows.append(fixed_row)
            
        # ЕСЛИ КОЛОНОК НЕ ХВАТАЕТ
        elif len(row) < expected_len:
            row.extend([''] * (expected_len - len(row)))
            fixed_rows.append(row)
            
        else:
            fixed_rows.append(row)
            
    return pd.DataFrame(fixed_rows, columns=header)


def standardize_dataframe(df):
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

    # Форматирование валюты
    if 'currency' in new_df.columns:
        new_df['currency'] = new_df['currency'].apply(normalize_currency)

    # Дефолтные значения (отель, адрес, картинка)
    mask_hotel = new_df['hotel_name'].isna() & new_df['id'].notna() & (new_df['id'] != '')
    clean_ids = new_df['id'].fillna('').astype(str).str.replace(r'\.0$', '', regex=True)
    new_df.loc[mask_hotel, 'hotel_name'] = "Hotel confirmation for reservation " + clean_ids[mask_hotel]

    mask_addr = new_df['address'].isna()
    new_df.loc[mask_addr, 'address'] = "You need to confirm your booking. This is required for verification purposes."

    mask_img = new_df['image'].isna()
    new_df.loc[mask_img, 'image'] = "https://i.ibb.co/C5dHd4fv/image.png"

    return new_df

def save_excel_perfect(df, filename):
    """Сохранение XLSX с выделенной шапкой и широкими колонками (защита от наложения)."""
    writer = pd.ExcelWriter(filename, engine='openpyxl')
    df.to_excel(writer, index=False, sheet_name='Data')
    worksheet = writer.sheets['Data']
    
    for col_letter, width in COLUMN_WIDTHS.items():
        worksheet.column_dimensions[col_letter].width = width
        
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
    
    msg = await update.message.reply_text("Обрабатываю файлы (чиню сдвиги колонок и конвертирую валюты)...")

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
            
            # Сохранение с указанием количества строк
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
