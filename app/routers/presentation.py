"""
Excel Commander - Presentation Router
Endpoints for PowerPoint generation.
"""
import os
from fastapi import APIRouter, HTTPException
from fastapi.responses import FileResponse
from app.models.schemas import PresentationRequest, PresentationResponse
from app.services.ai_service import get_ai_service
from app.services.pptx_service import get_pptx_service

router = APIRouter(prefix="/api/presentation", tags=["Presentation"])


@router.post("/generate", response_model=PresentationResponse)
async def generate_presentation(request: PresentationRequest):
    """
    Generate a PowerPoint presentation from Excel data.
    
    This endpoint:
    1. Analyzes the data using AI to generate insights
    2. Creates a professional PPTX with title, insights, chart, and table slides
    3. Returns a download URL for the generated file
    """
    ai_service = get_ai_service()
    pptx_service = get_pptx_service()
    
    try:
        # Validate data
        if len(request.data) < 2:
            return PresentationResponse(
                success=False,
                error="Veri en az 2 satır içermelidir (başlık + veri)."
            )
        
        # Generate insights using AI
        insights = ai_service.generate_insights(
            data=request.data,
            count=request.insights_count
        )
        
        # Map chart type
        chart_type_map = {
            "chart_bar": "bar",
            "chart_line": "line",
            "chart_pie": "pie"
        }
        chart_type = chart_type_map.get(request.chart_type.value, "bar")
        
        # Generate PPTX
        filepath = pptx_service.create_presentation(
            data=request.data,
            title=request.title,
            insights=insights,
            include_chart=request.include_chart,
            chart_type=chart_type
        )
        
        # Convert to relative URL for download
        filename = os.path.basename(filepath)
        download_url = f"/api/presentation/download/{filename}"
        
        return PresentationResponse(
            success=True,
            file_url=download_url,
            insights=insights
        )
        
    except Exception as e:
        return PresentationResponse(
            success=False,
            error=str(e)
        )


@router.get("/download/{filename}")
async def download_presentation(filename: str):
    """
    Download a generated PowerPoint file.
    """
    # Security: Only allow downloading from generated folder
    generated_dir = os.path.join(
        os.path.dirname(os.path.dirname(os.path.dirname(__file__))),
        "generated"
    )
    filepath = os.path.join(generated_dir, filename)
    
    # Validate file exists and is in allowed directory
    if not os.path.exists(filepath):
        raise HTTPException(status_code=404, detail="Dosya bulunamadı.")
    
    if not filepath.startswith(generated_dir):
        raise HTTPException(status_code=403, detail="Yetkisiz erişim.")
    
    return FileResponse(
        path=filepath,
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
