from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
from fastapi.middleware.cors import CORSMiddleware
import os
from dotenv import load_dotenv

load_dotenv()

app = FastAPI(title="Excel Commander API")

# Allow CORS for Office Add-in (localhost:3000 usually)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Production'da spesifik domain olacak
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

class FormulaRequest(BaseModel):
    description: str

@app.get("/")
def read_root():
    return {"status": "Excel Commander API is Online"}

@app.post("/generate-formula")
def generate_formula(req: FormulaRequest):
    # Mock response for now
    return {"formula": f"=SUM(A1:A10) # Mock for {req.description}"}

from fastapi.staticfiles import StaticFiles
# Mount frontend directory to serve the add-in files
# Assuming main.py is in /backend, we go up one level to /frontend
frontend_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "frontend")
if os.path.exists(frontend_path):
    app.mount("/taskpane", StaticFiles(directory=frontend_path, html=True), name="frontend")
else:
    print(f"Warning: Frontend path {frontend_path} not found")

if __name__ == "__main__":
    import uvicorn
    # Office Add-ins require HTTPS usually, but for localhost dev HTTP is sometimes okay or we need SSL.
    # We will stick to HTTP for local dev and assume user has certificates or uses a tunnel if needed.
    # To run: uvicorn main:app --reload
    uvicorn.run(app, host="0.0.0.0", port=8000)
