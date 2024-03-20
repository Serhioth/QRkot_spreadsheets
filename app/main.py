from fastapi import FastAPI
import uvicorn

from app.core.config import settings
from app.core.init_db import create_first_superuser
from app.api.routers import main_router


app = FastAPI(title=settings.app_title)
app.include_router(main_router)


@app.on_event('startup')
async def startup():
    await create_first_superuser()


if __name__ == '__main__':
    uvicorn.run('app.main:app', reload=True)
