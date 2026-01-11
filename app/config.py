"""
Excel Commander - Configuration
Handles environment variables and application settings.
"""
import os
from pydantic_settings import BaseSettings
from functools import lru_cache


class Settings(BaseSettings):
    """Application settings loaded from environment variables."""
    
    # API Keys
    openai_api_key: str = ""
    
    # Server Settings
    host: str = "0.0.0.0"
    port: int = 8000
    debug: bool = True
    
    # CORS Settings
    cors_origins: list[str] = ["*"]
    
    # AI Settings
    ai_model: str = "gpt-4o-mini"
    ai_temperature: float = 0.3  # Lower for more deterministic outputs
    ai_max_tokens: int = 1000
    
    class Config:
        env_file = ".env"
        env_file_encoding = "utf-8"


@lru_cache()
def get_settings() -> Settings:
    """Cached settings instance."""
    return Settings()
