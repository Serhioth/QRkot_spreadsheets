from datetime import datetime
from typing import Optional

from pydantic import (BaseModel,
                      Field,
                      Extra,
                      PositiveInt,
                      validator)

from app.constants import DEFAULT_INVESTED_AMOUNT, NAME_MIN_LENGTH, NAME_MAX_LENGTH


class CharityProjectBase(BaseModel):
    """Charity project base schema."""
    name: Optional[str] = Field(
        None,
        title='Название проекта',
        min_length=NAME_MIN_LENGTH,
        max_length=NAME_MAX_LENGTH
    )
    description: Optional[str] = Field(
        None,
        title='Описание проекта'
    )
    full_amount: Optional[PositiveInt] = Field(
        None,
        title='Требуемая сумма'
    )

    class Config:
        title = 'Базовая схема проекта'


class CharityProjectCreate(CharityProjectBase):
    """Charity project creation schema."""
    name: str = Field(
        ...,
        title='Название проекта',
        min_length=NAME_MIN_LENGTH,
        max_length=NAME_MAX_LENGTH
    )
    full_amount: PositiveInt = Field(
        ...,
        title='Требуемая сумма'
    )

    class Config:
        title = 'Схема для создания проекта'
        extra = Extra.forbid


class CharityProjectDB(CharityProjectCreate):
    """Charity project partially representation schema."""
    id: int = Field(
        ...,
        title='ID проекта'
    )
    invested_amount: int = Field(
        DEFAULT_INVESTED_AMOUNT,
        title='Собранная сумма'
    )
    fully_invested: bool = Field(
        default=False,
        title='Средства собраны'
    )
    create_date: datetime = Field(
        ...,
        title='Дата открытия проекта'
    )
    close_date: Optional[datetime] = Field(
        ...,
        title='Дата закрытия проекта'
    )

    class Config:
        title = 'Схема получения проекта'
        orm_mode = True
        schema_extra = {
            'example': {
                'id': 1,
                'name': 'Проект помощи полярным утконосам',
                'description': 'Мало того, что им холодно, так ещё и утиный клюв, вместо носа',
                'full_amount': 999,
                'invested_amount': 0,
                'fully_invested': 0,
                'create_date': '2024-03-01T02:18:40.662286'
            }
        }


class CharityProjectUpdate(CharityProjectBase):
    """Charity project partially update schema."""

    class Config:
        title = 'Схема для частичного обновления проекта'
        orm_mode = True
        extra = Extra.forbid
        schema_extra = {
            'example': {
                'name': 'Пандам на бамбук',
                'description': 'Голодные',
                'full_amount': '2500'
            }
        }

    @validator('name')
    def name_cant_be_null(cls, value: str):
        if value is None:
            raise ValueError(
                'Название проекта не должно быть пустым!'
            )
        return value

    @validator('description')
    def description_cant_be_null(cls, value: str):
        if value is None:
            raise ValueError(
                'Описание проекта не может быть пустым!'
            )
        return value
