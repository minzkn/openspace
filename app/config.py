from pydantic_settings import BaseSettings
from functools import lru_cache


class Settings(BaseSettings):
    secret_key: str = "0" * 64
    kek_key: str = "1" * 64
    database_url: str = "sqlite:///./openspace.db"
    session_ttl_seconds: int = 28800
    debug: bool = True
    max_upload_size: int = 52428800  # 50MB
    allowed_hosts: list[str] = ["*"]

    class Config:
        env_file = ".env"
        env_file_encoding = "utf-8"


@lru_cache
def get_settings() -> Settings:
    return Settings()


settings = get_settings()
