from http import HTTPStatus
from typing import Optional

from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.exc import IntegrityError
from sqlalchemy.ext.asyncio import AsyncSession

from app.core.db import get_async_session
from app.core.user import current_user, current_superuser
from app.crud.donation import donation_crud
from app.models.charity_project import CharityProject
from app.models.user import User
from app.schemas.donation import (
    DonationCreate,
    DonationDB,
    DonationDBShort
)
from app.services.invests import get_oldest_unclosed, make_invest


router = APIRouter()


@router.post(
    '/',
    response_model=DonationDBShort,
    response_model_exclude_none=True
)
async def create_donation(
    obj_in: DonationCreate,
    session: AsyncSession = Depends(
        get_async_session
    ),
    user: User = Depends(
        current_user
    )
):
    """Endpoint for creating new donations, registred user only."""

    new_donation = await donation_crud.create(
        obj_in,
        session,
        user
    )

    unclosed_projects = await get_oldest_unclosed(
        CharityProject,
        session
    )

    try:
        await make_invest(
            obj_to_invest=new_donation,
            unclosed_objects=unclosed_projects,
            session=session
        )

    except IntegrityError:
        await session.rollback()
        raise HTTPException(
            status_code=HTTPStatus.UNPROCESSABLE_ENTITY,
            detail='Нет проектов для распределения'
        )

    return new_donation


@router.get(
    '/',
    response_model=list[DonationDB],
    response_model_exclude_none=True,
    dependencies=[
        Depends(current_superuser)
    ]
)
async def get_all_donations(
    session: AsyncSession = Depends(
        get_async_session
    )
):
    """Endpoint for getting list of all donations, superuser only."""

    all_donations = await donation_crud.get_multi(
        session
    )
    return all_donations


@router.get(
    '/my',
    response_model=Optional[list[DonationDBShort]],
    response_model_exclude_none=True,
    response_model_exclude={'user_id'}
)
async def get_user_donations(
    user: User = Depends(
        current_user
    ),
    session: AsyncSession = Depends(
        get_async_session
    )
):
    """Endpoint for getting own investments, registred user only."""

    user_donations = await donation_crud.get_user_donations(
        user.id,
        session
    )
    return user_donations
