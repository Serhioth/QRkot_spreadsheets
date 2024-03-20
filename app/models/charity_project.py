from sqlalchemy import Column, String, Text

from app.models.base import DBObject
from app.models.constraints import (check_description_length,
                                    check_full_amount_is_positive)


class CharityProject(DBObject):
    """Charity project model."""

    name = Column(
        String(100),
        unique=True,
        nullable=False
    )
    description = Column(
        Text,
        nullable=False
    )

    __table_args__ = (
        check_description_length,
        check_full_amount_is_positive
    )

    def __repr__(self):
        return (
            f'Проект - {self.name},\n'
            f'Собрано - {self.invested_amount},\n'
            f'Необходимо собрать - {self.full_amount}'
        )
