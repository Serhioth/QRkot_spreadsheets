from http import HTTPStatus

from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.exc import IntegrityError
from sqlalchemy.ext.asyncio import AsyncSession

from app.api.validators import (
    charity_project_exists,
    check_description,
    check_name_duplicate,
    check_new_amount_greater_then_invested,
    check_project_is_closed,
    check_project_is_invested
)
from app.core.db import get_async_session
from app.core.user import current_superuser
from app.crud.charity_project import charity_project_crud
from app.models.donation import Donation
from app.schemas.charity_project import (
    CharityProjectCreate,
    CharityProjectDB,
    CharityProjectUpdate
)
from app.services.invests import get_oldest_unclosed, make_invest


router = APIRouter()


@router.post(
    '/',
    response_model=CharityProjectDB,
    response_model_exclude_none=True,
    dependencies=[Depends(current_superuser)]
)
async def create_new_charity_project(
    charity_project: CharityProjectCreate,
    session: AsyncSession = Depends(
        get_async_session
    )
):
    """Endpoint for creating charity projects, superuser only."""

    await check_name_duplicate(
        charity_project.name,
        session
    )
    check_description(charity_project.description)

    new_project = await charity_project_crud.create(
        charity_project,
        session
    )
    unclosed_donations = await get_oldest_unclosed(
        Donation,
        session
    )
    try:
        await make_invest(
            obj_to_invest=new_project,
            unclosed_objects=unclosed_donations,
            session=session
        )

    except IntegrityError:
        await session.rollback()
        raise HTTPException(
            status_code=HTTPStatus.UNPROCESSABLE_ENTITY,
            detail='Нет средств для распределения'
        )

    return new_project


@router.get(
    '/',
    response_model=list[CharityProjectDB],
    response_model_exclude_none=True
)
async def get_all_projects(
    session: AsyncSession = Depends(
        get_async_session
    )
):
    """Endpoint to get list of all charity projects."""

    all_projects = await charity_project_crud.get_multi(
        session
    )
    return all_projects


@router.patch(
    '/{charity_project_id}',
    response_model=CharityProjectDB,
    dependencies=[Depends(current_superuser)]
)
async def partially_update_charity_project(
    charity_project_id: int,
    obj_in: CharityProjectUpdate,
    session: AsyncSession = Depends(
        get_async_session
    )
):
    """Endpoint for updating charity projects, superuser only."""
    charity_project = await charity_project_exists(
        charity_project_id,
        session
    )

    check_project_is_closed(charity_project.fully_invested)

    if obj_in.name is not None:
        await check_name_duplicate(
            obj_in.name,
            session
        )

    if obj_in.description is not None:
        check_description(obj_in.description)

    if obj_in.full_amount is not None:
        check_new_amount_greater_then_invested(
            obj_in.full_amount,
            charity_project.invested_amount
        )

    await charity_project_crud.update(
        charity_project,
        obj_in,
        session
    )
    return charity_project


@router.delete(
    '/{charity_project_id}',
    response_model=CharityProjectDB,
    dependencies=[Depends(current_superuser)]
)
async def remove_charity_project(
    charity_project_id: int,
    session: AsyncSession = Depends(
        get_async_session
    )
):
    """Endpoint for removing charity projects, superuser only."""

    charity_project = await charity_project_exists(
        charity_project_id,
        session
    )
    check_project_is_invested(charity_project.invested_amount)

    charity_project = await charity_project_crud.remove(
        charity_project,
        session
    )
    return charity_project
