from datetime import datetime
from typing import Optional

from pydantic import BaseModel, Field, Extra, PositiveInt, validator

from app.constants import DEFAULT_INVESTED_AMOUNT


class DonationBase(BaseModel):
    """Donation base schema."""

    full_amount: PositiveInt = Field(
        ...,
        title='Сумма пожертвования'
    )
    comment: Optional[str] = Field(
        None,
        title='Комментарий'
    )


class DonationCreate(DonationBase):
    """Donation create schema."""

    class Config:
        extra = Extra.forbid
        title = 'Схема для создания пожертвования'
        schema_extra = {
            'example': {
                'full_amount': 999,
                'comment': 'На благое дело!'
            }
        }

    @validator('full_amount')
    def check_full_amount_greater_than_zero(
        cls, value: int
    ):
        if value == 0:
            raise ValueError(
                'Сумма пожертвования должна '
                'быть больше ноля!'
            )
        return value


class DonationDBShort(DonationBase):
    """Short object representation for regular users."""

    id: int = Field(
        ...,
        title='ID пожертвования'
    )
    create_date: datetime = Field(
        ...,
        title='Дата внесения'
    )

    class Config:
        title = 'Схема пожертвования для обычного пользователя'
        orm_mode = True
        schema_extra = {
            'example': {
                'full_amount': 999,
                'comment': 'На здоровье',
                'id': 1,
                'create_date': datetime.now()
            }
        }


class DonationDB(DonationDBShort):
    """Full object representation for superusers."""

    user_id: Optional[int] = Field(
        None,
        title='ID пожертвовавшего'
    )
    invested_amount: int = Field(
        DEFAULT_INVESTED_AMOUNT,
        title='Израсходовано'
    )
    fully_invested: bool = Field(
        False,
        title='Средства израсходованы'
    )
    close_date: Optional[datetime] = Field(
        None,
        title='Дата совершения пожертвования'
    )

    class Config:
        title = 'Схема пожертвования для администратора'
        orm_mode = True
        schema_extra = {
            'example': {
                'comment': '',
                'full_amount': 0,
                'id': 1,
                'create_date': datetime.now(),
                'user_id': 1,
                'invested_amount': 0,
                'fully_invested': 0
            }
        }
