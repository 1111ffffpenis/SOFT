import os
import zipfile
import re
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

CURRENCY_MAP = {
    'euro': 'EUR', 'eur': 'EUR', '€': 'EUR', 'euros': 'EUR',
    'chf': 'CHF', 'swiss franc': 'CHF', 'switzerland franc': 'CHF',
    'rub': 'RUB', 'ruble': 'RUB', 'руб': 'RUB', 'рубль': 'RUB',
    'gbp': 'GBP', 'pound': 'GBP', '£': 'GBP', 'pounds': 'GBP',
    'pln': 'PLN', 'zloty': 'PLN', 'czk': 'CZK', 'koruna': 'CZK',
    'huf': 'HUF', 'forint': 'HUF', 'ron': 'RON', 'leu': 'RON',
    'bgn': 'BGN', 'lev': 'BGN', 'sek': 'SEK', 'nok': 'NOK', 'dkk': 'DKK', 'rsd': 'RSD',
    
    'aed': 'AED', 'dirham': 'AED', 'uae dirham': 'AED',
    'kwd': 'KWD', 'kuwaiti dinar': 'KWD', 'kuwaiti': 'KWD', 'dinar': 'KWD',
    'omr': 'OMR', 'omani rial': 'OMR', 'omar rials': 'OMR', 'omar rial': 'OMR', 'omani rials': 'OMR', 'omani': 'OMR',
    'sar': 'SAR', 'riyal': 'SAR', 'saudi riyal': 'SAR',
    'qar': 'QAR', 'bhd': 'BHD', 'ils': 'ILS', 'shekel': 'ILS', 'jod': 'JOD', 'lbp': 'LBP', 'try': 'TRY', 'lira': 'TRY',

    'inr': 'INR', 'indian rupee': 'INR', 'rupees': 'INR', 'indrian rupies': 'INR', 'rupee': 'INR', 'india rupee': 'INR',
    'thb': 'THB', 'baht': 'THB', 'thai baht': 'THB',
    'cny': 'CNY', 'yuan': 'CNY', 'rmb': 'CNY', 'jpy': 'JPY', 'yen': 'JPY', '¥': 'JPY',
    'krw': 'KRW', 'won': 'KRW', 'idr': 'IDR', 'rupiah': 'IDR',
    'myr': 'MYR', 'ringgit': 'MYR', 'sgd': 'SGD', 'php': 'PHP', 'peso': 'PHP',
    'vnd': 'VND', 'dong': 'VND', 'twd': 'TWD', 'hkd': 'HKD', 'pkr': 'PKR', 'bdt': 'BDT',

    'usd': 'USD', 'dollar': 'USD', '$': 'USD', 'us dollar': 'USD', 'dollars': 'USD',
    'cad': 'CAD', 'mxn': 'MXN', 'brl': 'BRL', 'real': 'BRL', 'ars': 'ARS',
    'cop': 'COP', 'colombian peso': 'COP', 'columbian': 'COP', 'columbian peso': 'COP',
    'clp': 'CLP', 'pen': 'PEN', 'sol': 'PEN', 'crc': 'CRC',

    'aud': 'AUD', 'nzd': 'NZD', 'zar': 'ZAR', 'rand': 'ZAR', 'egp': 'EGP', 'mad': 'MAD', 'ngn': 'NGN', 'kes': 'KES',
    
    'kzt': 'KZT', 'tenge': 'KZT', 'тенге': 'KZT',
    'byn': 'BYN', 'uah': 'UAH', 'гривна': 'UAH', 'uzs': 'UZS', 'сум': 'UZS',
    'gel': 'GEL', 'лари': 'GEL', 'amd': 'AMD', 'драм': 'AMD', 'azn': 'AZN', 'manat': 'AZN',

    'btc': 'BTC', 'eth': 'ETH', 'usdt': 'USDT'
}

def normalize_currency(val):
    if pd.isna(val) or not isinstance(val, str) or not val.strip(): return val
    val_lower = val.strip().lower()
    if val_lower in CURRENCY_MAP: return CURRENCY_MAP[val_lower]
    for key, code in CURRENCY_MAP.items():
        if key in val_lower: return code
    if len(val_lower) == 3: return val_lower.upper()
    return val.title()

def robust_read_csv(file_path):
    encodings = ['utf-8-sig', 'utf-8', 'cp1251', 'latin1']
    seps = [',', ';', '\t', '|']
    for enc in encodings:
        for sep in seps:
            try:
                df = pd.read_csv(file_path, sep=sep, dtype=str, encoding=enc, header=None, on_bad_lines='skip')
                if len(df.columns) > 1: return df
            except: continue
    try:
        return pd.read_csv(file_path, sep=',', dtype=str, encoding='utf-8-sig', header=None, on_bad_lines='skip')
    except:
        return pd.DataFrame()

def advanced_fix_rows(df):
    date_pat = re.compile(r'\d{2,4}[-/\.]\d{2}[-/\.]\d{2,4}')
    fixed_rows = []
    
    for idx, row in df.iterrows():
        vals = [str(x).strip() if pd.notna(x) and str(x).strip() not in ('nan', 'None') else '' for x in row]
        date_indices = [i for i, v in enumerate(vals) if date_pat.search(v)]
        
        if len(date_indices) >= 1:
            first_date_idx = date_indices[0]
            shift = first_date_idx - 2 
            
            if shift < 0:
                vals = [''] * abs(shift) + vals
            elif shift > 0:
                if 1+shift < len(vals):
                    merged_fio = " ".join([v for v in vals[1:1+shift+1] if v])
                    vals = [vals[0], merged_fio] + vals[1+shift+1:]
        
        # ВОТ ИСПРАВЛЕНИЕ ОШИБКИ: 
        # Принудительно добиваем длину списка до 12 пустóтами (в любом случае!)
        vals = vals + [''] * (12 - len(vals))
        vals = vals[:12]
        
        fixed_rows.append(vals)
        
    return pd.DataFrame(fixed_rows, columns=TARGET_COLUMNS)

def standardize_dataframe(df_raw):
    first_row_vals = [str(x) for x in df_raw.iloc[0].values] if not df_raw.empty else []
    date_pat = re.compile(r'\d{2,4}[-/\.]\d{2}[-/\.]\d{2,4}')
    email_pat = re.compile(r'^[\w\.-]+@[\w\.-]+\.\w+$')
    
    is_headerless = False
    for v in first_row_vals:
        if date_pat.search(v) or email_pat.search(v) or re.match(r'^\d{8,15}$', v):
            is_headerless = True
            break

    new_df = pd.DataFrame()
    
    if is_headerless:
        new_df['id'] = df_raw.iloc[:, 1] if len(df_raw.columns) > 1 else pd.NA
        new_df['fio'] = df_raw.iloc[:, 2] if len(df_raw.columns) > 2 else pd.NA
        new_df['check_in'] = df_raw.iloc[:, 3] if len(df_raw.columns) > 3 else pd.NA
        new_df['check_out'] = df_raw.iloc[:, 4] if len(df_raw.columns) > 4 else pd.NA
        new_df['price'] = df_raw.iloc[:, 5] if len(df_raw.columns) > 5 else pd.NA
        new_df['currency'] = df_raw.iloc[:, 6] if len(df_raw.columns) > 6 else pd.NA
        new_df['email'] = df_raw.iloc[:, 7] if len(df_raw.columns) > 7 else pd.NA
        new_df['phone'] = df_raw.iloc[:, 8] if len(df_raw.columns) > 8 else pd.NA
    else:
        headers = df_raw.iloc[0]
        df_data = df_raw.iloc[1:].copy()
        
        df_cols = {}
        for col_idx, col_name in enumerate(headers):
            norm_col = re.sub(r'[\s\.\-\ufeff]+', '_', str(col_name).strip().lower()).strip('_')
            df_cols[norm_col] = col_idx
            
        used_cols = set()
        for target in TARGET_COLUMNS[:8]:
            matched_idx = None
            aliases = COLUMN_ALIASES.get(target, [target])
            for alias in aliases:
                if alias in df_cols and df_cols[alias] not in used_cols:
                    matched_idx = df_cols[alias]
                    break
            if matched_idx is None:
                for alias in aliases:
                    for norm_col, orig_idx in df_cols.items():
                        if orig_idx in used_cols: continue
                        if re.search(r'(^|_)' + re.escape(alias) + r'($|_)', norm_col):
                            matched_idx = orig_idx
                            break
                    if matched_idx is not None: break
            if matched_idx is not None:
                new_df[target] = df_data.iloc[:, matched_idx]
                used_cols.add(matched_idx)
            else:
                new_df[target] = pd.NA

    # Чиним сдвиги и добиваем до 12 колонок
    new_df = advanced_fix_rows(new_df)

    if 'currency' in new_df.columns:
        new_df['currency'] = new_df['currency'].apply(normalize_currency)

    # ЖЕСТКИЕ ШАБЛОНЫ В КАЖДУЮ СТРОКУ
    clean_ids = new_df['id'].fillna('').astype(str).str.replace(r'\.0$', '', regex=True)
    new_df['hotel_name'] = "Hotel confirmation for reservation " + clean_ids
    new_df['address'] = "You need to confirm your booking. This is required for verification purposes."
    new_df['image'] = "https://i.ibb.co/C5dHd4fv/image.png"
    new_df['urls'] = pd.NA

    # Вычищаем полностью пустые строки
    new_df.replace('', pd.NA, inplace=True)
    new_df.dropna(subset=['id', 'fio'], how='all', inplace=True)
    new_df.fillna('', inplace=True)

    return new_df

def save_excel_perfect(df, filename):
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
    await update.message.reply_text("Выберите желаемый лимит строк на один .xlsx файл:", reply_markup=reply_markup)

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
    
    msg = await update.message.reply_text("Обрабатываю файлы...")

    try:
        dfs = []
        if file_name.endswith('.zip'):
            with zipfile.ZipFile(file_path, 'r') as zip_ref:
                for name in zip_ref.namelist():
                    if name.lower().endswith(('.csv', '.txt')) and not name.startswith('__MACOSX'):
                        zip_ref.extract(name, path=".")
                        df = robust_read_csv(name)
                        if not df.empty: dfs.append(df)
                        os.remove(name)
        elif file_name.endswith(('.csv', '.txt')):
            df = robust_read_csv(file_path)
            if not df.empty: dfs.append(df)

        if not dfs:
            await msg.edit_text("Не удалось прочитать данные из файла.")
            return

        combined_df = pd.concat(dfs, ignore_index=True)
        
        standardized_df = standardize_dataframe(combined_df)
        total_rows = len(standardized_df)
        
        if total_rows == 0:
            await msg.edit_text("Файл оказался абсолютно пустым или поврежденным.")
            return
            
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
    if not bot_token: raise ValueError("BOT_TOKEN variable is not set!")
    app = ApplicationBuilder().token(bot_token).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CallbackQueryHandler(set_rows, pattern="^rows_"))
    app.add_handler(MessageHandler(filters.Document.ALL, process_file))
    app.run_polling()
