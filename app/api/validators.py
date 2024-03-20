from http import HTTPStatus
from typing import Optional

from fastapi import HTTPException
from sqlalchemy.ext.asyncio import AsyncSession

from app.crud.charity_project import charity_project_crud


def check_new_amount_greater_then_invested(
    new_amount: int,
    invested_amount: int
):
    """Checks new amount value."""

    if new_amount < invested_amount:
        raise HTTPException(
            status_code=400,
            detail=(
                'Новая сумма сбора не может быть меньше '
                'суммы уже собранных средств!'
            )
        )


async def charity_project_exists(
    charity_project_id: int,
    session: AsyncSession
):
    """Checks, if charity project already exists."""

    charity_project = await charity_project_crud.get(
        charity_project_id,
        session
    )
    if charity_project is None:
        raise HTTPException(
            status_code=HTTPStatus.NOT_FOUND,
            detail='Проект не найден!'
        )
    return charity_project


async def check_name_duplicate(
    project_name: str,
    session: AsyncSession
):
    """Checks projects name duplications."""

    charity_project_id = await charity_project_crud.get_project_id_by_name(
        project_name,
        session
    )
    if charity_project_id is not None:
        raise HTTPException(
            status_code=HTTPStatus.BAD_REQUEST,
            detail='Проект с таким именем уже существует!'
        )


def check_description(description: Optional[str]):
    """Checks, if description is correct."""

    if description is None or len(description) == 0:
        raise HTTPException(
            status_code=HTTPStatus.UNPROCESSABLE_ENTITY,
            detail='Добавьте описание к проекту'
        )


def check_project_is_invested(invested_amount: int):
    """Checking that no funds have been contributed to the project."""

    if invested_amount > 0:
        raise HTTPException(
            status_code=HTTPStatus.BAD_REQUEST,
            detail='В проект были внесены средства, не подлежит удалению!'
        )


def check_project_is_closed(investment_status: bool):
    """Checking that project is not closed before editing."""

    if investment_status:
        raise HTTPException(
            status_code=HTTPStatus.BAD_REQUEST,
            detail='Закрытый проект нельзя редактировать!'
        )
