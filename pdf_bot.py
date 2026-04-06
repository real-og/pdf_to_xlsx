from aiogram import types
from aiogram import executor
from aiogram import Bot, Dispatcher
from aiogram.contrib.fsm_storage.memory import MemoryStorage
import logging
from aiogram.utils.exceptions import FileIsTooBig
from config import BOT_TOKEN
import parser
import adding_price
import exchange_rate
import os


def delete_file(file_name):
    file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), file_name)

    if os.path.exists(file_path):
        os.remove(file_path)


logging.basicConfig(level=logging.WARNING)


INPUT_PDF = "waybill.pdf"
OUTPUT_XLSX = "waybill.xlsx"

storage = MemoryStorage()

bot = Bot(token=BOT_TOKEN, parse_mode="HTML")
dp = Dispatcher(bot, storage=storage)

@dp.message_handler(commands=['start'], state="*")
async def send_welcome(message: types.Message):
    await message.answer('pong')


@dp.message_handler(state="*", content_types=types.ContentType.ANY)
async def send_welcome(mes: types.Message):
    if mes.content_type in ['document']:
        try:
            await mes.document.download(destination_file=mes.document.file_name)

            if '.xlsx' in mes.document.file_name:
                file_name_first = mes.document.file_name
                file_name_second = mes.document.file_name[:-5] + '_modified.xlsx'
                adding_price.enrich_ozon_prices_xlsx(mes.document.file_name, exchange_rate.get_exchange_rate())
                await mes.answer_document(types.InputFile(file_name_second))
                delete_file(file_name_second)
                delete_file(file_name_first)
            else:

                parser.parse_pdf_to_xlsx(mes.document.file_name, mes.document.file_name.replace('.pdf', '.xlsx'))
                await mes.answer_document(types.InputFile(mes.document.file_name.replace('.pdf', '.xlsx')))

        except FileIsTooBig:
            
            await mes.reply('файл больше 20мб')
    else:
        await mes.reply('это не док')



if __name__ == '__main__':
    print("Starting parsing pdf to xlsx bot")
    executor.start_polling(dp, skip_updates=True)