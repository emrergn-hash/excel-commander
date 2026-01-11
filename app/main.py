"""
Excel Commander - Main Application Entry Point
A professional Excel AI Assistant API.
"""
import os
import logging
from contextlib import asynccontextmanager

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles

from app.config import get_settings
from app.models.schemas import HealthResponse
from app.routers import formula, presentation
from app.services.ai_service import get_ai_service

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)


@asynccontextmanager
async def lifespan(app: FastAPI):
    """Application lifespan events."""
    # Startup
    settings = get_settings()
    logger.info("🚀 Excel Commander API starting...")
    logger.info(f"   AI Model: {settings.ai_model}")
    logger.info(f"   Debug Mode: {settings.debug}")
    
    # Initialize services
    ai_service = get_ai_service()
    if ai_service.is_configured():
        logger.info("   AI Service: ✅ Configured")
    else:
        logger.warning("   AI Service: ⚠️ Not configured (using mocks)")
    
    yield
    
    # Shutdown
    logger.info("👋 Excel Commander API shutting down...")


# Create FastAPI app
app = FastAPI(
    title="Excel Commander API",
    description="AI-powered Excel Assistant - Generate formulas, create presentations, and clean data.",
    version="1.0.0",
    lifespan=lifespan
)

# CORS Middleware
settings = get_settings()
app.add_middleware(
    CORSMiddleware,
    allow_origins=settings.cors_origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Include Routers
app.include_router(formula.router)
app.include_router(presentation.router)

# Mount frontend static files
frontend_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "frontend")
if os.path.exists(frontend_path):
    app.mount("/taskpane", StaticFiles(directory=frontend_path, html=True), name="frontend")
    logger.info(f"   Frontend: ✅ Mounted at /taskpane")
else:
    logger.warning(f"   Frontend: ⚠️ Path not found: {frontend_path}")

# Mount generated files for download
generated_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "generated")
os.makedirs(generated_path, exist_ok=True)
app.mount("/generated", StaticFiles(directory=generated_path), name="generated")


# ============ Root Endpoints ============

@app.get("/", response_model=HealthResponse, tags=["Health"])
async def health_check():
    """
    Health check endpoint.
    Returns API status and configuration.
    """
    ai_service = get_ai_service()
    return HealthResponse(
        status="online",
        version="1.0.0",
        ai_configured=ai_service.is_configured()
    )


@app.get("/api", tags=["Health"])
async def api_info():
    """
    API information endpoint.
    """
    return {
        "name": "Excel Commander API",
        "version": "1.0.0",
        "endpoints": {
            "formula": {
                "generate": "POST /api/formula/generate",
                "explain": "POST /api/formula/explain",
                "clean": "POST /api/formula/clean"
            },
            "presentation": {
                "generate": "POST /api/presentation/generate",
                "download": "GET /api/presentation/download/{filename}"
            }
        }
    }


# ============ Development Server ============

if __name__ == "__main__":
    import uvicorn
    settings = get_settings()
    uvicorn.run(
        "app.main:app",
        host=settings.host,
        port=settings.port,
        reload=settings.debug
    )
