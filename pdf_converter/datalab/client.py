"""
Datalab API Client for PDF to Markdown conversion.
Uses the Marker API for high-accuracy OCR of financial documents.
"""

import os
import time
from pathlib import Path
from typing import Optional
from dataclasses import dataclass
import httpx
from rich.console import Console
from rich.progress import Progress, SpinnerColumn, TextColumn

console = Console()


@dataclass
class DatalabResult:
    """Result from Datalab API processing."""
    success: bool
    markdown: Optional[str] = None
    html: Optional[str] = None
    page_count: int = 0
    error: Optional[str] = None
    cost_breakdown: Optional[dict] = None
    runtime: float = 0.0


class DatalabClient:
    """
    Datalab API client for document conversion.
    
    Uses the Marker API endpoint to convert PDFs to structured formats
    with high accuracy OCR suitable for financial documents.
    
    Pricing:
    - $25 free credit per month
    - Fast mode: $4/1000 pages
    - Balanced mode: $4/1000 pages
    - Accurate mode: $6/1000 pages
    """
    
    BASE_URL = "https://www.datalab.to/api/v1"
    
    def __init__(
        self,
        api_key: Optional[str] = None,
        mode: str = "balanced",  # "fast", "balanced", "accurate"
        output_format: str = "markdown",  # "markdown", "html", "json", "chunks"
        poll_interval: float = 2.0,
        max_wait_time: float = 600.0,  # 10 minutes max
        verify_ssl: bool = False  # Disable for corporate proxies
    ):
        self.api_key = api_key or os.environ.get("DATALAB_API_KEY", "").strip()
        self.mode = mode
        self.output_format = output_format
        self.poll_interval = poll_interval
        self.max_wait_time = max_wait_time
        self.verify_ssl = verify_ssl
        
        if not self.api_key:
            console.print("[yellow]⚠️ No DATALAB_API_KEY found. Set it in .env or pass directly.[/yellow]")
        
        # Initialize HTTP client with SSL config
        self._client = httpx.Client(
            verify=self.verify_ssl,
            timeout=httpx.Timeout(60.0, connect=30.0)
        )
    
    def _get_headers(self) -> dict:
        """Get headers for API requests."""
        return {
            "X-API-Key": self.api_key,
            "Accept": "application/json"
        }
    
    def convert_pdf(
        self,
        pdf_path: str,
        mode: Optional[str] = None,
        output_format: Optional[str] = None,
        page_range: Optional[str] = None,
        paginate: bool = True
    ) -> DatalabResult:
        """
        Convert a PDF file to markdown/html using Datalab Marker API.
        
        Args:
            pdf_path: Path to the PDF file
            mode: Processing mode ("fast", "balanced", "accurate")
            output_format: Output format ("markdown", "html", "json")
            page_range: Specific pages to process (e.g., "0,2-4,6")
            paginate: Add page separators to output
        
        Returns:
            DatalabResult with converted content
        """
        if not self.api_key:
            return DatalabResult(
                success=False,
                error="No API key configured. Set DATALAB_API_KEY environment variable."
            )
        
        pdf_path = Path(pdf_path)
        if not pdf_path.exists():
            return DatalabResult(
                success=False,
                error=f"PDF file not found: {pdf_path}"
            )
        
        mode = mode or self.mode
        output_format = output_format or self.output_format
        
        console.print(f"[cyan]📤 Uploading PDF to Datalab ({mode} mode)...[/cyan]")
        
        # Step 1: Submit PDF for processing
        try:
            with open(pdf_path, "rb") as f:
                files = {"file": (pdf_path.name, f, "application/pdf")}
                data = {
                    "mode": mode,
                    "output_format": output_format,
                    "paginate": str(paginate).lower()
                }
                if page_range:
                    data["page_range"] = page_range
                
                response = self._client.post(
                    f"{self.BASE_URL}/marker",
                    files=files,
                    data=data,
                    headers=self._get_headers()
                )
                response.raise_for_status()
        except httpx.HTTPStatusError as e:
            return DatalabResult(
                success=False,
                error=f"Upload failed: {e.response.status_code} - {e.response.text}"
            )
        except Exception as e:
            return DatalabResult(
                success=False,
                error=f"Upload error: {str(e)}"
            )
        
        result = response.json()
        
        if not result.get("success", False):
            return DatalabResult(
                success=False,
                error=result.get("error", "Unknown error during submission")
            )
        
        request_id = result.get("request_id")
        check_url = result.get("request_check_url")
        
        console.print(f"[green]✓ PDF submitted. Request ID: {request_id}[/green]")
        
        # Step 2: Poll for results
        return self._poll_for_result(request_id, check_url)
    
    def _poll_for_result(self, request_id: str, check_url: Optional[str] = None) -> DatalabResult:
        """
        Poll the API until processing is complete.
        
        Args:
            request_id: The request ID from submission
            check_url: Optional direct URL to check status
        
        Returns:
            DatalabResult with the converted content
        """
        url = check_url or f"{self.BASE_URL}/marker/{request_id}"
        start_time = time.time()
        
        with Progress(
            SpinnerColumn(),
            TextColumn("[progress.description]{task.description}"),
            console=console
        ) as progress:
            task = progress.add_task("Processing PDF...", total=None)
            
            while (time.time() - start_time) < self.max_wait_time:
                try:
                    response = self._client.get(url, headers=self._get_headers())
                    response.raise_for_status()
                    result = response.json()
                except httpx.HTTPStatusError as e:
                    return DatalabResult(
                        success=False,
                        error=f"Status check failed: {e.response.status_code}"
                    )
                except Exception as e:
                    return DatalabResult(
                        success=False,
                        error=f"Status check error: {str(e)}"
                    )
                
                status = result.get("status", "").lower()
                
                if status == "complete":
                    progress.update(task, description="[green]✓ Processing complete![/green]")
                    
                    return DatalabResult(
                        success=result.get("success", True),
                        markdown=result.get("markdown"),
                        html=result.get("html"),
                        page_count=result.get("page_count", 0),
                        cost_breakdown=result.get("cost_breakdown"),
                        runtime=result.get("runtime", time.time() - start_time)
                    )
                
                elif status in ["failed", "error"]:
                    return DatalabResult(
                        success=False,
                        error=result.get("error", "Processing failed")
                    )
                
                # Still processing
                elapsed = time.time() - start_time
                progress.update(task, description=f"Processing... ({elapsed:.0f}s elapsed)")
                time.sleep(self.poll_interval)
        
        return DatalabResult(
            success=False,
            error=f"Timeout after {self.max_wait_time} seconds"
        )
    
    def check_health(self) -> bool:
        """
        Check if the Datalab API is available.
        
        Returns:
            True if API is healthy, False otherwise
        """
        try:
            response = self._client.get(
                f"{self.BASE_URL}/user_health",
                headers=self._get_headers()
            )
            if response.status_code == 200:
                result = response.json()
                return result.get("status") == "ok"
        except Exception:
            pass
        return False
    
    def close(self):
        """Close the HTTP client."""
        self._client.close()
    
    def __enter__(self):
        return self
    
    def __exit__(self, *args):
        self.close()
