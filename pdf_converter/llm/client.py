"""
LLM Client for AI-powered data extraction.
Supports Anthropic Claude with chunking strategy to avoid token truncation.
"""

import json
import os
import re
from typing import Optional, Any
from dataclasses import dataclass
from rich.console import Console

console = Console()


@dataclass
class ExtractionResult:
    """Result from LLM extraction."""
    success: bool
    data: dict
    raw_response: str
    error: Optional[str] = None
    tokens_used: int = 0


class LLMClient:
    """
    LLM Client with chunking support for large PDF extraction.
    Uses conservative token limits to avoid output truncation.
    Supports both Anthropic Claude and OpenAI GPT.
    """
    
    def __init__(
        self,
        model: str = "claude-sonnet-4-20250514",
        max_tokens_output: int = 16000,
        temperature: float = 0.0,
        max_retries: int = 3,
        provider: str = "auto"  # "anthropic", "openai", or "auto"
    ):
        self.model = model
        self.max_tokens_output = max_tokens_output
        self.temperature = temperature
        self.max_retries = max_retries
        self.provider = provider
        self._client = None
        self._provider_type = None
        self._init_client()
    
    def _init_client(self):
        """Initialize the LLM client (Anthropic, OpenAI, or Gemini)."""
        import httpx
        
        # Load .env first if available
        try:
            from dotenv import load_dotenv
            load_dotenv()
        except ImportError:
            pass
        
        # Try Anthropic first
        if self.provider in ["auto", "anthropic"]:
            try:
                from anthropic import Anthropic
                api_key = os.environ.get("ANTHROPIC_API_KEY", "").strip()
                
                if api_key:
                    http_client = httpx.Client(verify=False)
                    self._client = Anthropic(api_key=api_key, http_client=http_client)
                    self._provider_type = "anthropic"
                    console.print("[dim]Using Anthropic Claude[/dim]")
                    return
            except ImportError:
                pass
        
        # Try OpenAI as fallback
        if self.provider in ["auto", "openai"]:
            try:
                from openai import OpenAI
                api_key = os.environ.get("OPENAI_API_KEY", "").strip()
                
                if api_key:
                    http_client = httpx.Client(verify=False)
                    self._client = OpenAI(api_key=api_key, http_client=http_client)
                    self._provider_type = "openai"
                    self.model = "gpt-4o"  # Use GPT-4o for OpenAI
                    console.print("[dim]Using OpenAI GPT-4o[/dim]")
                    return
            except ImportError:
                pass
        
        # Try Google Gemini as fallback (using direct HTTP calls)
        if self.provider in ["auto", "gemini"]:
            api_key = os.environ.get("GEMINI_API_KEY", "").strip()
            
            if api_key:
                # Use direct httpx calls instead of google-genai library
                # This gives us full control over SSL verification
                self._gemini_api_key = api_key
                self._client = httpx.Client(verify=False, timeout=600.0)
                self._provider_type = "gemini_direct"
                self.model = "gemini-2.0-flash"
                console.print("[dim]Using Google Gemini 2.0 Flash (Direct HTTP)[/dim]")
                return
        
        console.print("[yellow]⚠️ No API key found (ANTHROPIC_API_KEY, OPENAI_API_KEY, or GEMINI_API_KEY). LLM extraction will be simulated.[/yellow]")
    
    def extract(
        self,
        prompt: str,
        system_prompt: Optional[str] = None,
        expected_keys: Optional[list] = None
    ) -> ExtractionResult:
        """
        Extract data using LLM with JSON schema enforcement.
        
        Args:
            prompt: The extraction prompt with PDF text
            system_prompt: Optional system prompt
            expected_keys: List of expected keys in JSON response
        
        Returns:
            ExtractionResult with parsed data
        """
        if not self._client:
            return self._simulate_extraction(prompt, expected_keys)
        
        default_system = """You are a precise data extraction assistant. 
Return ONLY valid JSON. No markdown code blocks. No explanations. No additional text.
If a cell is empty, use 0 for numbers and "" for text. NEVER use null.
Extract ALL data from ALL pages provided. Do not stop early."""
        
        system = system_prompt or default_system
        
        for attempt in range(self.max_retries):
            try:
                if self._provider_type == "anthropic":
                    response = self._client.messages.create(
                        model=self.model,
                        max_tokens=self.max_tokens_output,
                        temperature=self.temperature,
                        system=system,
                        messages=[{"role": "user", "content": prompt}]
                    )
                    raw_text = response.content[0].text
                    tokens_used = response.usage.output_tokens if hasattr(response, 'usage') else 0
                elif self._provider_type == "openai":
                    response = self._client.chat.completions.create(
                        model=self.model,
                        max_tokens=self.max_tokens_output,
                        temperature=self.temperature,
                        messages=[
                            {"role": "system", "content": system},
                            {"role": "user", "content": prompt}
                        ]
                    )
                    raw_text = response.choices[0].message.content
                    tokens_used = response.usage.completion_tokens if hasattr(response, 'usage') else 0
                elif self._provider_type == "gemini_direct":
                    # Direct HTTP call to Gemini API
                    full_prompt = f"{system}\n\n{prompt}"
                    url = f"https://generativelanguage.googleapis.com/v1beta/models/{self.model}:generateContent?key={self._gemini_api_key}"
                    payload = {
                        "contents": [{"parts": [{"text": full_prompt}]}],
                        "generationConfig": {
                            "temperature": self.temperature,
                            "maxOutputTokens": self.max_tokens_output,
                        }
                    }
                    response = self._client.post(url, json=payload)
                    response.raise_for_status()
                    result = response.json()
                    raw_text = result["candidates"][0]["content"]["parts"][0]["text"]
                    tokens_used = result.get("usageMetadata", {}).get("candidatesTokenCount", 0)
                else:  # gemini (google-genai package)
                    full_prompt = f"{system}\n\n{prompt}"
                    response = self._client.models.generate_content(
                        model=self.model,
                        contents=full_prompt,
                        config={
                            "temperature": self.temperature,
                            "max_output_tokens": self.max_tokens_output,
                        }
                    )
                    raw_text = response.text
                    tokens_used = 0
                
                # Parse JSON
                data = self._parse_json(raw_text)
                
                # Validate expected keys
                if expected_keys:
                    missing = [k for k in expected_keys if k not in data]
                    if missing:
                        console.print(f"[yellow]Warning: Missing keys: {missing}[/yellow]")
                
                return ExtractionResult(
                    success=True,
                    data=data,
                    raw_response=raw_text,
                    tokens_used=tokens_used
                )
                
            except json.JSONDecodeError as e:
                console.print(f"[yellow]Attempt {attempt + 1}: JSON parse error: {e}[/yellow]")
                if attempt < self.max_retries - 1:
                    # Try to repair JSON
                    prompt = f"""The previous response had invalid JSON. 
Please fix and return ONLY valid JSON:

{raw_text}

Error: {e}

Return corrected JSON only:"""
                else:
                    return ExtractionResult(
                        success=False,
                        data={},
                        raw_response=raw_text if 'raw_text' in locals() else "",
                        error=f"JSON parse error after {self.max_retries} attempts: {e}"
                    )
                    
            except Exception as e:
                console.print(f"[red]LLM error: {e}[/red]")
                if attempt == self.max_retries - 1:
                    return ExtractionResult(
                        success=False,
                        data={},
                        raw_response="",
                        error=str(e)
                    )
        
        return ExtractionResult(success=False, data={}, raw_response="", error="Unknown error")
    
    def _parse_json(self, text: str) -> dict:
        """Parse JSON with common fixes for LLM output."""
        text = text.strip()
        
        # Remove markdown code blocks
        if text.startswith("```"):
            lines = text.split("\n")
            # Remove first and last lines if they're code block markers
            if lines[0].startswith("```"):
                lines = lines[1:]
            if lines and lines[-1].strip() == "```":
                lines = lines[:-1]
            text = "\n".join(lines)
        
        # Try direct parse
        try:
            return json.loads(text)
        except json.JSONDecodeError:
            pass
        
        # Try to extract JSON object/array
        json_match = re.search(r'[\[{].*[\]}]', text, re.DOTALL)
        if json_match:
            try:
                return json.loads(json_match.group())
            except json.JSONDecodeError:
                pass
        
        # Try to balance braces/brackets
        text = self._balance_json(text)
        return json.loads(text)
    
    def _balance_json(self, text: str) -> str:
        """Attempt to balance JSON braces and brackets."""
        # Count brackets
        open_braces = text.count('{')
        close_braces = text.count('}')
        open_brackets = text.count('[')
        close_brackets = text.count(']')
        
        # Add missing closers
        text += '}' * (open_braces - close_braces)
        text += ']' * (open_brackets - close_brackets)
        
        return text
    
    def _simulate_extraction(self, prompt: str, expected_keys: Optional[list]) -> ExtractionResult:
        """Simulate extraction when no LLM client is available."""
        console.print("[yellow]🔄 Simulating LLM extraction (no API key)[/yellow]")
        
        # Return empty structure with expected keys
        data = {}
        if expected_keys:
            for key in expected_keys:
                data[key] = []
        
        return ExtractionResult(
            success=True,
            data=data,
            raw_response="{}",
            error="Simulated extraction - no API key"
        )
    
    def extract_with_continuation(
        self,
        prompt: str,
        continuation_context: str = "",
        system_prompt: Optional[str] = None,
        expected_keys: Optional[list] = None
    ) -> ExtractionResult:
        """
        Extract with continuation context for multi-chunk processing.
        Injects context from previous chunks into the prompt.
        """
        if continuation_context:
            full_prompt = f"""{continuation_context}

{prompt}"""
        else:
            full_prompt = prompt
        
        return self.extract(full_prompt, system_prompt, expected_keys)


class ChunkedExtractor:
    """
    Handles extraction of large PDFs by processing in chunks.
    Maintains context continuity across chunks.
    """
    
    def __init__(
        self,
        llm_client: LLMClient,
        max_pages_per_chunk: int = 5,
        overlap_pages: int = 1
    ):
        self.llm = llm_client
        self.max_pages_per_chunk = max_pages_per_chunk
        self.overlap_pages = overlap_pages
    
    def extract_section(
        self,
        pdf_reader,
        start_page: int,
        end_page: int,
        prompt_template: str,
        section_key: str,
        context_builder: callable = None
    ) -> list:
        """
        Extract a section spanning multiple pages with chunking.
        
        Args:
            pdf_reader: PDFReader instance
            start_page: Start page (0-indexed)
            end_page: End page (0-indexed, inclusive)
            prompt_template: Prompt template with {text} placeholder
            section_key: Key to extract from JSON response
            context_builder: Optional function to build continuation context
        
        Returns:
            List of all extracted rows
        """
        all_rows = []
        context = ""
        
        current_page = start_page
        while current_page <= end_page:
            chunk_end = min(current_page + self.max_pages_per_chunk - 1, end_page)
            
            console.print(f"  [dim]Processing pages {current_page + 1}-{chunk_end + 1}...[/dim]")
            
            # Extract text for this chunk
            text = pdf_reader.extract_pages_text(current_page, chunk_end)
            
            # Build prompt
            prompt = prompt_template.format(text=text)
            
            # Add continuation context if available
            if context:
                prompt = f"{context}\n\n{prompt}"
            
            # Extract
            result = self.llm.extract(prompt, expected_keys=[section_key])
            
            if result.success and section_key in result.data:
                rows = result.data[section_key]
                
                # Update context for next chunk
                if context_builder and rows:
                    context = context_builder(rows)
                
                all_rows.extend(rows)
            else:
                console.print(f"[yellow]Warning: Failed to extract chunk pages {current_page + 1}-{chunk_end + 1}[/yellow]")
            
            # Move to next chunk (with overlap)
            current_page = chunk_end + 1 - self.overlap_pages
            if current_page <= chunk_end:
                current_page = chunk_end + 1
        
        return all_rows
