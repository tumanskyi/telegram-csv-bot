import os
import pandas as pd
from datetime import datetime
from telegram import Update
from telegram.ext import ApplicationBuilder, MessageHandler, filters, ContextTypes

BOT_TOKEN = '7920319045:AAGSImgpTFi8Jr4Aa3Xf9KF8wwzoXVtm_e8'

async def handle_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    file = update.message.document
    if not file.file_name.endswith('.csv'):
        await update.message.reply_text("Пожалуйста, загрузите CSV-файл.")
        return

    # Скачать файл
    file_path = f"{file.file_id}.csv"
    new_file = await file.get_file()
    await new_file.download_to_drive(file_path)

    # Обработка CSV
    df = pd.read_csv(file_path)

    # 1. Вкладка Заявки
    df_zayavki = df.copy()

    # 2. Вкладка Доставка курьером
    df_delivery = df[df['Delivery'].str.contains("Доставка курьером", na=False)]

    # 3. Вкладка Заявки для поставщика — группировка по товарам
    all_items = []
    for items in df['Товары в заказе'].dropna():
        for line in str(items).split('\n'):
            name_qty = line.split('x')
            if len(name_qty) == 2:
                name = name_qty[0].strip()
                qty = int(name_qty[1].split('≡')[0].strip())
                all_items.append((name, qty))
            else:
                all_items.append((line.strip(), 1))

    summary_df = pd.DataFrame(all_items, columns=["Товар", "Количество"])
    summary_df = summary_df.groupby("Товар", as_index=False).sum()

    # Сохранение Excel
    out_path = f"Заявки_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df_zayavki.to_excel(writer, sheet_name="Заявки", index=False)
        df_delivery.to_excel(writer, sheet_name="Доставка курьером", index=False)
        summary_df.to_excel(writer, sheet_name="Заявки для поставщика", index=False)

    await update.message.reply_document(document=open(out_path, "rb"))

    os.remove(file_path)
    os.remove(out_path)

app = ApplicationBuilder().token(BOT_TOKEN).build()
app.add_handler(MessageHandler(filters.Document.ALL, handle_file))
app.run_polling()
