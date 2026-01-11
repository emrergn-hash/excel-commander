"""
Excel Commander - Formula Router
Endpoints for formula generation and explanation.
"""
from fastapi import APIRouter, HTTPException
from app.models.schemas import (
    FormulaRequest, FormulaResponse,
    ExplainRequest, ExplainResponse,
    CleanDataRequest, CleanDataResponse
)
from app.services.ai_service import get_ai_service

router = APIRouter(prefix="/api/formula", tags=["Formula"])


@router.post("/generate", response_model=FormulaResponse)
async def generate_formula(request: FormulaRequest):
    """
    Generate an Excel formula from natural language description.
    
    Example:
        Input: "A sütunundaki toplam satışları hesapla"
        Output: "=TOPLA(A:A)"
    """
    ai_service = get_ai_service()
    
    try:
        formula, explanation = ai_service.generate_formula(
            description=request.description,
            context=request.context
        )
        
        if formula is None:
            return FormulaResponse(
                success=False,
                error="Formül oluşturulamadı."
            )
        
        return FormulaResponse(
            success=True,
            formula=formula,
            explanation=explanation
        )
        
    except Exception as e:
        return FormulaResponse(
            success=False,
            error=str(e)
        )


@router.post("/explain", response_model=ExplainResponse)
async def explain_formula(request: ExplainRequest):
    """
    Explain an Excel formula in simple terms.
    
    Example:
        Input: "=DÜŞEYARA(A1;B:C;2;0)"
        Output: "Bu formül A1 hücresindeki değeri B:C aralığında arar..."
    """
    ai_service = get_ai_service()
    
    try:
        explanation = ai_service.explain_formula(request.formula)
        
        return ExplainResponse(
            success=True,
            explanation=explanation
        )
        
    except Exception as e:
        return ExplainResponse(
            success=False,
            error=str(e)
        )


@router.post("/clean", response_model=CleanDataResponse)
async def clean_data(request: CleanDataRequest):
    """
    Clean and standardize data.
    
    Operations:
    - Trim whitespace
    - Standardize capitalization
    - Fix date formats
    - Remove duplicates
    """
    try:
        cleaned = []
        changes = []
        
        for row_idx, row in enumerate(request.data):
            cleaned_row = []
            for col_idx, cell in enumerate(row):
                original = cell
                
                if isinstance(cell, str):
                    # Strip whitespace
                    new_val = cell.strip()
                    
                    # Title case for names (if looks like a name)
                    if new_val and new_val[0].islower():
                        new_val = new_val.title()
                    
                    if new_val != original:
                        changes.append(f"[{row_idx+1},{col_idx+1}]: '{original}' → '{new_val}'")
                    
                    cleaned_row.append(new_val)
                else:
                    cleaned_row.append(cell)
            
            cleaned.append(cleaned_row)
        
        return CleanDataResponse(
            success=True,
            cleaned_data=cleaned,
            changes_made=changes[:10]  # Limit to 10 changes in response
        )
        
    except Exception as e:
        return CleanDataResponse(
            success=False,
            error=str(e)
        )
