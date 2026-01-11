"""
Excel Commander - AI Service
Handles all AI API interactions via OpenRouter.
OpenRouter provides access to multiple AI models including free options.
"""
import logging
from typing import Optional, List, Any
import httpx
from app.config import get_settings

logger = logging.getLogger(__name__)

# OpenRouter API Base URL
OPENROUTER_BASE_URL = "https://openrouter.ai/api/v1"


class AIService:
    """Service class for AI operations using OpenRouter."""
    
    SYSTEM_PROMPT_FORMULA = """Sen bir Excel formül uzmanısın. Kullanıcının isteğine göre doğru Excel formülünü üret.
Kurallar:
1. SADECE formülü döndür, açıklama ekleme.
2. Formül Excel syntax'ına tam uymalı (Türkçe Excel için noktalı virgül kullan).
3. Formül daima '=' ile başlamalı.
4. Geçersiz istek gelirse "HATA: [sebep]" döndür.
"""

    SYSTEM_PROMPT_EXPLAIN = """Sen bir Excel eğitmenisin. Verilen formülü adım adım açıkla.
Kurallar:
1. Türkçe açıkla.
2. Teknik jargon kullanma, basit dilde anlat.
3. Maddeler halinde açıkla.
"""

    SYSTEM_PROMPT_INSIGHTS = """Sen bir veri analistisin. Verilen tabloyu analiz et ve önemli içgörüler (insights) çıkar.
Kurallar:
1. Türkçe yaz.
2. Kısa ve öz maddeler halinde yaz.
3. Sayısal değerlere atıfta bulun.
4. İş kararlarına yardımcı olacak yorumlar yap.
"""

    # Free models on OpenRouter (as of 2025)
    FREE_MODELS = [
        "meta-llama/llama-3.2-3b-instruct:free",
        "google/gemma-2-9b-it:free",
        "mistralai/mistral-7b-instruct:free",
        "qwen/qwen-2-7b-instruct:free",
    ]

    def __init__(self):
        settings = get_settings()
        self.api_key = settings.openai_api_key  # Using same env var for simplicity
        self.model = settings.ai_model
        self.temperature = settings.ai_temperature
        self.max_tokens = settings.ai_max_tokens
        
        # Check if we should use a free model
        if self.model == "gpt-4o-mini" and self.api_key.startswith("sk-or-"):
            # Default to a good free model on OpenRouter
            self.model = "meta-llama/llama-3.2-3b-instruct:free"
            logger.info(f"Using free OpenRouter model: {self.model}")

    def is_configured(self) -> bool:
        """Check if AI service is properly configured."""
        return bool(self.api_key)

    def _call_openrouter(self, messages: List[dict], max_tokens: int = None) -> Optional[str]:
        """Make a call to OpenRouter API."""
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
            "HTTP-Referer": "https://excel-commander.app",  # Required by OpenRouter
            "X-Title": "Excel Commander"
        }
        
        payload = {
            "model": self.model,
            "messages": messages,
            "temperature": self.temperature,
            "max_tokens": max_tokens or self.max_tokens
        }
        
        try:
            with httpx.Client(timeout=30.0) as client:
                response = client.post(
                    f"{OPENROUTER_BASE_URL}/chat/completions",
                    headers=headers,
                    json=payload
                )
                response.raise_for_status()
                data = response.json()
                return data["choices"][0]["message"]["content"].strip()
        except httpx.HTTPStatusError as e:
            logger.error(f"OpenRouter API error: {e.response.status_code} - {e.response.text}")
            return None
        except Exception as e:
            logger.error(f"OpenRouter call failed: {e}")
            return None

    def generate_formula(self, description: str, context: Optional[str] = None) -> tuple[str, str]:
        """
        Generate an Excel formula based on user description.
        Returns: (formula, explanation)
        """
        if not self.is_configured():
            return self._mock_formula(description)
        
        prompt = f"Kullanıcı İsteği: {description}"
        if context:
            prompt += f"\nBağlam: {context}"
        
        messages = [
            {"role": "system", "content": self.SYSTEM_PROMPT_FORMULA},
            {"role": "user", "content": prompt}
        ]
        
        formula = self._call_openrouter(messages)
        
        if formula is None:
            return self._mock_formula(description)
        
        # Validate formula starts with '=' or is an error
        if not formula.startswith("=") and not formula.startswith("HATA"):
            formula = "=" + formula
        
        # Get explanation
        explanation = self._explain_formula(formula)
        
        return formula, explanation

    def explain_formula(self, formula: str) -> str:
        """Explain an Excel formula in simple terms."""
        if not self.is_configured():
            return f"Bu formül ({formula}) verilerinizi hesaplar. (Mock açıklama)"
        
        return self._explain_formula(formula)

    def _explain_formula(self, formula: str) -> str:
        """Internal method to explain a formula."""
        messages = [
            {"role": "system", "content": self.SYSTEM_PROMPT_EXPLAIN},
            {"role": "user", "content": f"Bu formülü açıkla: {formula}"}
        ]
        
        result = self._call_openrouter(messages, max_tokens=500)
        return result or "Açıklama oluşturulamadı."

    def generate_insights(self, data: List[List[Any]], count: int = 3) -> List[str]:
        """
        Analyze data and generate business insights.
        """
        if not self.is_configured():
            return self._mock_insights(data, count)
        
        data_str = self._format_data_for_prompt(data)
        
        messages = [
            {"role": "system", "content": self.SYSTEM_PROMPT_INSIGHTS},
            {"role": "user", "content": f"Bu veriyi analiz et ve {count} adet içgörü çıkar:\n\n{data_str}"}
        ]
        
        result = self._call_openrouter(messages, max_tokens=800)
        
        if result is None:
            return self._mock_insights(data, count)
        
        # Split by newlines and filter empty lines
        insights = [line.strip() for line in result.split("\n") if line.strip()]
        return insights[:count]

    def _format_data_for_prompt(self, data: List[List[Any]]) -> str:
        """Format 2D data array for AI prompt."""
        if not data:
            return "Boş veri"
        
        lines = []
        for i, row in enumerate(data[:20]):
            lines.append(" | ".join(str(cell) for cell in row))
        return "\n".join(lines)

    def _mock_formula(self, description: str) -> tuple[str, str]:
        """Mock formula generation for testing."""
        desc_lower = description.lower()
        
        if "topla" in desc_lower or "sum" in desc_lower:
            return "=TOPLA(A1:A10)", "Bu formül A1'den A10'a kadar olan hücreleri toplar."
        elif "ortalama" in desc_lower or "average" in desc_lower:
            return "=ORTALAMA(A1:A10)", "Bu formül A1'den A10'a kadar olan değerlerin ortalamasını hesaplar."
        elif "say" in desc_lower or "count" in desc_lower:
            return "=BAĞ_DEĞ_SAY(A1:A10)", "Bu formül A1'den A10'a kadar dolu hücreleri sayar."
        elif "eğer" in desc_lower or "if" in desc_lower:
            return '=EĞER(A1>100;"Yüksek";"Düşük")', "Bu formül A1 100'den büyükse 'Yüksek', değilse 'Düşük' yazar."
        elif "düşeyara" in desc_lower or "vlookup" in desc_lower:
            return "=DÜŞEYARA(A1;Tablo!A:B;2;0)", "Bu formül A1 değerini Tablo'da arar ve 2. sütundaki karşılığını getirir."
        else:
            return f"=TOPLA(A:A)", f"'{description}' için örnek formül oluşturuldu."

    def _mock_insights(self, data: List[List[Any]], count: int) -> List[str]:
        """Mock insights for testing."""
        return [
            "📈 Veriler genel olarak yükseliş trendi gösteriyor.",
            "📊 En yüksek değer son satırlarda gözlemleniyor.",
            "💡 Büyüme oranı pozitif seyrediyor."
        ][:count]


# Singleton instance
_ai_service: Optional[AIService] = None

def get_ai_service() -> AIService:
    """Get or create AIService singleton."""
    global _ai_service
    if _ai_service is None:
        _ai_service = AIService()
    return _ai_service
