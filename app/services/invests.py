from datetime import datetime
from typing import Union, Optional

from sqlalchemy import select
from sqlalchemy.ext.asyncio import AsyncSession

from app.models.charity_project import CharityProject
from app.models.donation import Donation


async def get_oldest_unclosed(
    model: Union[CharityProject, Donation],
    session: AsyncSession
) -> Optional[Union[CharityProject, Donation]]:
    """Function for getting all unclosed objects of model."""
    db_obj = await session.execute(
        select(model).where(
            model.fully_invested == 0
        ).order_by(model.create_date)
    )
    return db_obj.scalars().all()


def close_object(
    obj: Union[CharityProject, Donation]
):
    """Function for closing fully invested objects."""
    obj.fully_invested = True
    obj.close_date = datetime.now()
    return obj


async def make_invest(
    obj_to_invest: Union[CharityProject, Donation],
    unclosed_objects: Optional[list[Union[CharityProject, Donation]]],
    session: AsyncSession
):
    """Function for investing into unclosed objects."""
    if unclosed_objects:
        for invest_object in unclosed_objects:
            invest_remainder = (
                obj_to_invest.full_amount - obj_to_invest.invested_amount
            )
            required_invest = (
                invest_object.full_amount - invest_object.invested_amount
            )
            if invest_remainder >= required_invest:
                obj_to_invest.invested_amount += required_invest
                invest_object.invested_amount = invest_object.full_amount
                close_object(invest_object)
                if invest_remainder == required_invest:
                    close_object(obj_to_invest)
            else:
                invest_object.invested_amount += invest_remainder
                obj_to_invest.invested_amount = obj_to_invest.full_amount
                close_object(obj_to_invest)
                break

        await session.commit()
        await session.refresh(obj_to_invest)

    return obj_to_invest
