import pandas as pd
import os
from aiogram import executor, Bot, Dispatcher, types
from aiogram.contrib.fsm_storage.memory import MemoryStorage
from aiogram.dispatcher.filters import Text
from aiogram.dispatcher.filters.state import StatesGroup, State
from aiogram.types import ReplyKeyboardMarkup, KeyboardButton
from aiogram.dispatcher import FSMContext
from dotenv import load_dotenv

load_dotenv()

TOKEN_API = os.getenv("TOKEN_API")


def analyz(dataframe, vibor):
    grupa = dataframe['Группа'] == vibor
    marks = dataframe['Оценка'].count()
    marks_grupa = dataframe['Группа'] == vibor
    kolvo_marks_grupa=dataframe[marks_grupa]['Оценка'].count()
    students_grupa = dataframe[marks_grupa]['Личный номер студента'].unique()
    formi_control = dataframe[marks_grupa]['Уровень контроля'].unique()
    Years = sorted(dataframe['Год'].unique())
    return(f"В исходном датасете содержалось,{marks}, оценок из них, {kolvo_marks_grupa}, оценок относятся к группе, {vibor},\n В датасете находятся оценки {len(students_grupa)} студентов со следующими личными номерами: {', '.join(map(str,students_grupa))} \n Используемые формы контроля: {', '.join(map(str,formi_control))}, \n Данные представлены по следующим учебным годам: {', '.join(map(str,Years))}")

storage = MemoryStorage()
bot = Bot(TOKEN_API)
dp = Dispatcher(bot=bot,
                storage=storage)


class ProfileStatesGroup(StatesGroup):
    Wait_For_Document = State()
    Choose_Group = State()


@dp.message_handler(commands=['cancel'], state='*')
async def cancel_handler(message: types.Message, state: FSMContext):
    if await state.get_state():
        await state.finish()
        await message.answer("Операция отменена.")
    else:
        await message.answer("Нет активных операций для отмены.")


@dp.message_handler(commands=['start'])
async def cmd_start(message: types.Message) -> None:
    await message.answer("Привет, для начала отправь мне файл excel")
    await ProfileStatesGroup.Wait_For_Document.set()


@dp.message_handler(state=ProfileStatesGroup.Wait_For_Document)
async def not_document(message:types.Message,state: FSMContext):
    if message.text == '/cancel':
        await state.finish()
        await message.answer("Операция отменена.")
    else:
        await message.answer("Пожалуйста, отправьте файл в формате Excel (.xlsx). Если у вас нет файла, используйте команду /cancel.")

@dp.message_handler(content_types=types.ContentType.DOCUMENT, state=ProfileStatesGroup.Wait_For_Document)

async def send_document(message: types.Message, state: FSMContext):
    await message.answer("Подождите, документ обрабатывается")
    # Получение информации о документе
    file_id = message.document.file_id
    file_info = await bot.get_file(file_id)
    file_path = file_info.file_path

    # Скачивание документа
    downloaded_file = await bot.download_file(file_path)
    file_extension = os.path.splitext(file_info.file_path)[1].lower()
    # Проверка что файл excel
    if message.document.mime_type == 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' and file_extension == '.xlsx':
        # Сохранение документа
        with open('user_uploaded_file.xlsx', 'wb') as new_file:
            new_file.write(downloaded_file.read())
        # Анализ файла и отправка результата
        dataframe = pd.read_excel('user_uploaded_file.xlsx')
        expected_columns = ["Группа", "Оценка","Год", "Личный номер студента", "Уровень контроля"]
        column = []
        if all(column in dataframe.columns for column in expected_columns):
            all_groups = dataframe['Группа'].unique()
            await message.answer("Файл успешно загружен и проанализирован.")
            await message.answer(f"Для начала анализа необходимо выбрать одну из групп:\n{', '.join(all_groups)}")
            await ProfileStatesGroup.Choose_Group.set()
        else:
            await message.answer("Пожалуйста, отправьте корректный файл, в котором присутствуют необходимые столбцы")
    else:
        await message.answer("Пожалуйста, отправьте файл в формате Excel (.xlsx).")

@dp.message_handler(state=ProfileStatesGroup.Choose_Group)
async def choose(message:types.message,state:FSMContext)-> None:
    dataframe = pd.read_excel('user_uploaded_file.xlsx')
    async with state.proxy() as data:
        data['group'] = message.text
    all_groups = dataframe['Группа'].unique()
    if data['group'] in all_groups:
        await message.answer(analyz(dataframe,data['group']))
        await state.finish()
    else:
        await message.answer("Пожалуйста, выберете группу из файла")

if __name__ == '__main__':
    executor.start_polling(dp,
                           skip_updates=True)