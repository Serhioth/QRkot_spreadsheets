from datetime import datetime

from aiogoogle import Aiogoogle

from app.core.config import settings


DATETIME_FORMAT = "%Y/%m/%d %H:%M:%S"


async def spreadsheets_create(wrapper_services: Aiogoogle) -> str:
    # Получаем текущую дату для заголовка документа
    now_date_time = datetime.now().strftime(DATETIME_FORMAT)
    # Создаём экземпляр класса Resource
    service = await wrapper_services.discover('sheets', 'v4')
    # Формируем тело запроса
    spreadsheet_body = {
        'properties': {'title': f'Отчёт от {now_date_time}',
                       'locale': 'ru_RU'},
        'sheets': [{'properties': {'sheetType': 'GRID',
                                   'sheetId': 0,
                                   'title': 'Лист1',
                                   'gridProperties': {'rowCount': 100,
                                                      'columnCount': 11}}}]
    }
    # Выполняем запрос
    response = await wrapper_services.as_service_account(
        service.spreadsheets.create(json=spreadsheet_body)
    )
    spreadsheetid = response['spreadsheetId']
    return spreadsheetid


async def set_user_permissions(
        spreadsheetid: str,
        wrapper_services: Aiogoogle
) -> None:
    permissions_body = {'type': 'user',
                        'role': 'writer',
                        'emailAddress': settings.email}
    service = await wrapper_services.discover('drive', 'v3')
    await wrapper_services.as_service_account(
        service.permissions.create(
            fileId=spreadsheetid,
            json=permissions_body,
            fields="id"
        ))


def format_duration(duration_sec: int) -> str:
    """Function to format duration to required format."""

    days, seconds = divmod(duration_sec, 86400)
    hours, seconds = divmod(seconds, 3600)
    minutes, seconds = divmod(seconds, 60)

    duration_str = f"{days} days, {hours:02}:{minutes:02}:{seconds:02}"

    return duration_str


async def spreadsheets_update_value(
        spreadsheetid: str,
        closed_projects: list,
        wrapper_services: Aiogoogle
) -> None:
    now_date_time = datetime.now().strftime(DATETIME_FORMAT)
    service = await wrapper_services.discover('sheets', 'v4')
    # Здесь формируется тело таблицы
    table_values = [
        ['Отчёт от', now_date_time],
        ['Топ проектов по скорости закрытия'],
        ['Название проекта', 'Время сбора', 'Описание']
    ]
    # Здесь в таблицу добавляются строчки
    for project, duration in closed_projects:
        new_row = [project.name, format_duration(duration), project.description]
        table_values.append(new_row)

    update_body = {
        'majorDimension': 'ROWS',
        'values': table_values
    }
    response = await wrapper_services.as_service_account(  # noqa
        service.spreadsheets.values.update(
            spreadsheetId=spreadsheetid,
            range='A1:E30',
            valueInputOption='USER_ENTERED',
            json=update_body
        )
    )
