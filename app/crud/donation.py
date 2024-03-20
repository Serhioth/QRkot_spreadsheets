from sqlalchemy import select
from sqlalchemy.ext.asyncio import AsyncSession

from app.crud.base import CRUDBase
from app.models.donation import Donation


class CRUDDonation(CRUDBase):
    """Donation CRUD class."""

    async def get_user_donations(
        self,
        user_id: int,
        session: AsyncSession
    ):
        user_donations = await session.execute(
            select(Donation).where(
                Donation.user_id == user_id
            )
        )
        return user_donations.scalars().all()


donation_crud = CRUDDonation(Donation)
