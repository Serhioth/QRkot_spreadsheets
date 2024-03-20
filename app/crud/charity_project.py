from typing import Optional

from sqlalchemy import extract, select
from sqlalchemy.ext.asyncio import AsyncSession

from app.crud.base import CRUDBase
from app.models.charity_project import CharityProject


class CRUDCharityProject(CRUDBase):
    """Charity project CRUD class."""

    async def get_project_id_by_name(
        self,
        project_name: str,
        session: AsyncSession
    ) -> Optional[int]:
        project_id = await session.execute(
            select(CharityProject.id).where(
                CharityProject.name == project_name
            )
        )
        project_id = project_id.scalars().first()

        return project_id

    async def get_projects_by_completion_rate(
        self,
        session: AsyncSession
    ) -> Optional[list[CharityProject]]:

        projects = await session.execute(
            select(
                CharityProject,
                (
                    extract('epoch', CharityProject.close_date) -
                    extract('epoch', CharityProject.create_date)
                ).label("duration")
            ).where(CharityProject.fully_invested).order_by("duration")
        )

        return projects.all()


charity_project_crud = CRUDCharityProject(CharityProject)
