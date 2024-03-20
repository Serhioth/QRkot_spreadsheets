from sqlalchemy import Column, ForeignKey, Text, Integer

from app.models.base import DBObject


class Donation(DBObject):
    """Donation model."""

    user_id = Column(
        Integer,
        ForeignKey('user.id', name='user_donation')
    )
    comment = Column(
        Text,
        nullable=True
    )

    def __repr__(self):
        return (
            f'Пожертвование на сумму {self.full_amount},\n'
            f'Потрачено - {self.invested_amount}, \n'
            f'Комментарий - {self.comment}'
        )
