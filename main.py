import os
import zipfile
import pandas as pd
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, CallbackQueryHandler, ContextTypes, filters

USER_SETTINGS = {}

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    keyboard = [
        [InlineKeyboardButton("1000", callback_data="rows_1000"), InlineKeyboardButton("2000", callback_data="rows_2000")],
        [InlineKeyboardButton("3000", callback_data="rows_3000"), InlineKeyboardButton("5000", callback_data="rows_5000")]
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)
    await update.message.reply_text("Выберите количество строк в одном .xlsx файле:", reply_markup=reply_markup)

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
    
    await update.message.reply_text("Обрабатываю данные...")

    try:
        dfs = []
        if file_name.endswith('.zip'):
            with zipfile.ZipFile(file_path, 'r') as zip_ref:
                for name in zip_ref.namelist():
                    if name.lower().endswith(('.csv', '.txt')) and not name.startswith('__MACOSX'):
                        with zip_ref.open(name) as f:
                            # dtype=str предотвращает потерю данных (например, нулей в начале телефонов)
                            df = pd.read_csv(f, sep=None, engine='python', dtype=str)
                            dfs.append(df)
        elif file_name.endswith(('.csv', '.txt')):
            df = pd.read_csv(file_path, sep=None, engine='python', dtype=str)
            dfs.append(df)

        if not dfs:
            await update.message.reply_text("Подходящие CSV/TXT файлы не найдены.")
            return

        combined_df = pd.concat(dfs, ignore_index=True)
        total_rows = len(combined_df)
        
        parts_count = (total_rows + chunk_size - 1) // chunk_size
        for part in range(parts_count):
            start_idx = part * chunk_size
            end_idx = start_idx + chunk_size
            chunk = combined_df.iloc[start_idx:end_idx]
            
            out_name = f"output_part_{part + 1}.xlsx"
            chunk.to_excel(out_name, index=False)
            
            with open(out_name, 'rb') as f:
                await update.message.reply_document(document=f, filename=out_name)
            os.remove(out_name)

        await update.message.reply_text(f"Готово! Обработано всего строк: {total_rows}")

    except Exception as e:
        await update.message.reply_text(f"Ошибка при обработке: {str(e)}")
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