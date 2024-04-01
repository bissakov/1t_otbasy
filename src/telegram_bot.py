import asyncio
import os
from io import BytesIO

import dotenv
from PIL import Image
from retry import retry
from telegram import Bot, InputFile
from telegram.error import TelegramError


class TelegramBot:
    def __init__(self, token: str, chat_id: str) -> None:
        self.bot = Bot(token)
        self.chat_id = chat_id

    @retry(TelegramError, tries=5, delay=2, backoff=2)
    async def send_message(self, message: str) -> None:
        await self.bot.send_message(chat_id=self.chat_id, text=message)

    @retry(TelegramError, tries=5, delay=2, backoff=2)
    async def send_picture(self, image: Image, caption: str = None) -> None:
        image_io = BytesIO()
        image.save(image_io, format='PNG')
        image_io.seek(0)
        await self.bot.send_photo(
            chat_id=self.chat_id,
            photo=InputFile(image_io, filename='image.png'),
            caption=caption)

    @retry(TelegramError, tries=5, delay=2, backoff=2)
    async def send_document(self, document_path: str, caption: str = None) -> None:
        await self.bot.send_document(chat_id=self.chat_id, document=open(document_path, 'rb'), caption=caption)


# dotenv.load_dotenv()
# BOT = TelegramBot(token=os.getenv('TOKEN'), chat_id=os.getenv('CHAT_ID'))
#
#
# def send_message(message: str) -> None:
#     loop = asyncio.new_event_loop()
#     asyncio.set_event_loop(loop)
#     loop.run_until_complete(BOT.send_message(message=message))
#     loop.close()
#
#
# def send_picture(image: Image, caption: str = None) -> None:
#     loop = asyncio.new_event_loop()
#     asyncio.set_event_loop(loop)
#     loop.run_until_complete(BOT.send_picture(image=image, caption=caption))
#     loop.close()
#
#
# def send_document(document_path: str, caption: str = None) -> None:
#     loop = asyncio.new_event_loop()
#     asyncio.set_event_loop(loop)
#     loop.run_until_complete(BOT.send_document(document_path=document_path, caption=caption))
#     loop.close()
