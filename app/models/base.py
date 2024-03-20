from datetime import datetime

from sqlalchemy import Column, Integer, Boolean, DateTime

from app.core.db import Base


class DBObject(Base):
    """Base abstract model."""

    __abstract__ = True

    full_amount = Column(
        Integer,
        nullable=False
    )
    invested_amount = Column(
        Integer,
        nullable=False,
        default=0
    )
    fully_invested = Column(
        Boolean,
        nullable=False,
        default=False
    )
    create_date = Column(
        DateTime,
        nullable=False,
        default=datetime.now
    )
    close_date = Column(
        DateTime
    )
