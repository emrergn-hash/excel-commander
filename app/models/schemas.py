"""
Excel Commander - Pydantic Models (Schemas)
Request and response models for API endpoints.
"""
from pydantic import BaseModel, Field
from typing import Optional, List, Any
from enum import Enum


# ============ Enums ============
class CommandType(str, Enum):
    """Types of commands the AI can handle."""
    FORMULA = "formula"
    EXPLAIN = "explain"
    CLEAN = "clean"
    PRESENTATION = "presentation"


class SlideLayout(str, Enum):
    """Available PowerPoint slide layouts."""
    TITLE = "title"
    BULLET_POINTS = "bullet_points"
    CHART_BAR = "chart_bar"
    CHART_LINE = "chart_line"
    CHART_PIE = "chart_pie"
    TWO_COLUMN = "two_column"


# ============ Request Models ============
class FormulaRequest(BaseModel):
    """Request model for formula generation."""
    description: str = Field(..., min_length=3, max_length=500)
    context: Optional[str] = Field(None, description="Optional context about the data")
    language: str = Field("tr", description="Language code (tr, en)")

    class Config:
        json_schema_extra = {
            "example": {
                "description": "A sütunundaki tüm sayıları topla",
                "context": "Satış verileri tablosu",
                "language": "tr"
            }
        }


class ExplainRequest(BaseModel):
    """Request model for formula explanation."""
    formula: str = Field(..., min_length=1)
    language: str = Field("tr")


class CleanDataRequest(BaseModel):
    """Request model for data cleaning."""
    data: List[List[Any]] = Field(..., description="2D array of cell values")
    instructions: Optional[str] = Field(None, description="Specific cleaning instructions")


class PresentationRequest(BaseModel):
    """Request model for PowerPoint generation."""
    data: List[List[Any]] = Field(..., description="2D array of cell values (with headers)")
    title: Optional[str] = Field("Analiz Raporu", description="Presentation title")
    insights_count: int = Field(3, ge=1, le=5, description="Number of insights to generate")
    include_chart: bool = Field(True)
    chart_type: SlideLayout = Field(SlideLayout.CHART_BAR)

    class Config:
        json_schema_extra = {
            "example": {
                "data": [
                    ["Ay", "Satış", "Kar"],
                    ["Ocak", 10000, 2000],
                    ["Şubat", 12000, 2500],
                    ["Mart", 15000, 3000]
                ],
                "title": "2025 Q1 Satış Raporu",
                "insights_count": 3,
                "include_chart": True,
                "chart_type": "chart_bar"
            }
        }


# ============ Response Models ============
class FormulaResponse(BaseModel):
    """Response model for formula generation."""
    success: bool
    formula: Optional[str] = None
    explanation: Optional[str] = None
    error: Optional[str] = None


class ExplainResponse(BaseModel):
    """Response model for formula explanation."""
    success: bool
    explanation: Optional[str] = None
    error: Optional[str] = None


class CleanDataResponse(BaseModel):
    """Response model for data cleaning."""
    success: bool
    cleaned_data: Optional[List[List[Any]]] = None
    changes_made: Optional[List[str]] = None
    error: Optional[str] = None


class PresentationResponse(BaseModel):
    """Response model for PowerPoint generation."""
    success: bool
    file_url: Optional[str] = None  # URL to download the generated PPTX
    insights: Optional[List[str]] = None
    error: Optional[str] = None


class HealthResponse(BaseModel):
    """Response model for health check."""
    status: str
    version: str
    ai_configured: bool
