import os
import pandas as pd
from datetime import datetime
from telegram import Update
from telegram.ext import ApplicationBuilder, MessageHandler, filters, ContextTypes

# Получаем токен из переменной окружения (как на Render)
BOT_TOKEN = os.environ.get("BOT_TOKEN")

async def handle_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        file = update.message.document
        if not file.file_name.endswith('.csv'):
            await update.message.reply_text("Пожалуйста, отправьте CSV-файл.")
            return

        file_path = f"{file.file_id}.csv"
        new_file = await file.get_file()
        await new_file.download_to_drive(file_path)

        try:
            df = pd.read_csv(file_path)
        except Exception as e:
            await update.message.reply_text(f"Ошибка при чтении CSV-файла: {e}")
            os.remove(file_path)
            return

        # Проверка нужных колонок
        required_columns = ['Delivery', 'Товары в заказе']
        missing = [col for col in required_columns if col not in df.columns]
        if missing:
            await update.message.reply_text(f"В файле отсутствуют обязательные колонки: {', '.join(missing)}")
            os.remove(file_path)
            return

        if df.empty:
            await update.message.reply_text("Файл пустой.")
            os.remove(file_path)
            return

        # Вкладка 1: Заявки
        df_zayavki = df.copy()

        # Вкладка 2: Доставка курьером
        df_delivery = df[df['Delivery'].str.contains("Доставка курьером", na=False)]

        # Вкладка 3: Заявки для поставщика
        all_items = []
        for items in df['Товары в заказе'].dropna():
            for line in str(items).split('\n'):
                parts = line.split('x')
                if len(parts) == 2:
                    name = parts[0].strip()
                    try:
                        qty = int(parts[1].split('≡')[0].strip())
                    except:
                        qty = 1
                    all_items.append((name, qty))
                else:
                    all_items.append((line.strip(), 1))

        if not all_items:
            await update.message.reply_text("Не удалось извлечь товары из заказов.")
            os.remove(file_path)
            return

        summary_df = pd.DataFrame(all_items, columns=["Товар", "Количество"])
        summary_df = summary_df.groupby("Товар", as_index=False).sum()

        # Сохранение Excel
        out_path = f"Zayavki_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
            df_zayavki.to_excel(writer, sheet_name="Заявки", index=False)
            df_delivery.to_excel(writer, sheet_name="Доставка курьером", index=False)
            summary_df.to_excel(writer, sheet_name="Заявки для поставщика", index=False)

        await update.message.reply_document(document=open(out_path, "rb"))

    except Exception as error:
        await update.message.reply_text(f"Произошла ошибка: {error}")

    finally:
        # Очистка файлов
        if os.path.exists(file_path):
            os.remove(file_path)
        if os.path.exists(out_path):
            os.remove(out_path)

# Запуск приложения
app = ApplicationBuilder().token(BOT_TOKEN).build()
app.add_handler(MessageHandler(filters.Document.ALL, handle_file))
app.run_polling()
