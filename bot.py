import logging
import pandas as pd
from telegram import Update
from telegram.ext import (
    ApplicationBuilder,
    MessageHandler,
    CommandHandler,
    ContextTypes,
    filters
)
import os

# ----------------------------------------
# PUT YOUR BOT TOKEN HERE
# ----------------------------------------
BOT_TOKEN = "YOUR_NEW_TOKEN_HERE"

# ----------------------------------------
# Logging
# ----------------------------------------
logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO
)

# ----------------------------------------
# Category Function
# ----------------------------------------
def categorize(name):
    n = str(name).lower()

    if "flower" in n:
        return "Flowers"

    fruits_kw = [
        "apple","banana","mango","papaya","pineapple","orange",
        "grape","watermelon","melon","amla","guava","pomegranate",
        "kiwi"
    ]
    if any(k in n for k in fruits_kw):
        return "Fruits"

    exotic_kw = [
        "avocado","zucchini","broccoli","asparagus","lettuce",
        "dragon","berry","cherry","leek","celery","blueberry"
    ]
    if any(k in n for k in exotic_kw):
        return "Exotic"

    return "Vegetables"


# ----------------------------------------
# Handle Incoming Excel File
# ----------------------------------------
async def handle_file(update: Update, context: ContextTypes.DEFAULT_TYPE):

    if not update.message.document:
        return

    file = await update.message.document.get_file()
    await file.download_to_drive("input.xlsx")

    # Read Excel
    df = pd.read_excel("input.xlsx")

    # Apply category
    df["Category"] = df["Product Description"].apply(categorize)

    # Sort by category
    df_sorted = df.sort_values("Category")

    # Save new file
    df_sorted.to_excel("output.xlsx", index=False)

    # Send processed file back
    await update.message.reply_document(
        document=open("output.xlsx", "rb"),
        caption="✔ Your categorized & sorted Excel is ready!"
    )


# ----------------------------------------
# /start message
# ----------------------------------------
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Send an Excel file (.xlsx). I will sort & categorize it automatically!"
    )


# ----------------------------------------
# Main Runner
# ----------------------------------------
def main():
    app = ApplicationBuilder().token(BOT_TOKEN).build()

    # Commands
    app.add_handler(CommandHandler("start", start))

    # Excel documents
    app.add_handler(MessageHandler(filters.Document.FileExtension("xlsx"), handle_file))

    # Optional: Handle all documents
    app.add_handler(MessageHandler(filters.Document.ALL, handle_file))

    app.run_polling()


if __name__ == "__main__":
    main()
