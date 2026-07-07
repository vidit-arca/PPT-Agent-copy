"""
Configuration management for PPT Agent.
Handles environment variables, LLM provider selection, and settings.
"""

import os
from typing import Literal
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()


class Config:
    """Application configuration."""
    
    def __init__(self):
        # LLM Provider Settings
        self.llm_provider: Literal["openai", "ollama", "azure", "aws"] = os.getenv("LLM_PROVIDER", "ollama").lower()
        self.llm_model: str = os.getenv("LLM_MODEL", self._get_default_model())
        self.llm_temperature: float = float(os.getenv("LLM_TEMPERATURE", "0"))
        
        # OpenAI Settings
        self.openai_api_key: str = os.getenv("OPENAI_API_KEY", "")
        self.openai_base_url: str = os.getenv("OPENAI_BASE_URL", "https://api.openai.com/v1")
        
        # Azure OpenAI Settings
        self.azure_api_key: str = os.getenv("AZURE_OPENAI_API_KEY", "")
        self.azure_endpoint: str = os.getenv("AZURE_OPENAI_ENDPOINT", "https://tejomaya.openai.azure.com/")
        self.azure_api_version: str = os.getenv("AZURE_API_VERSION", "2024-12-01-preview")
        self.azure_deployment: str = os.getenv("AZURE_DEPLOYMENT_NAME", "gpt-4o-mini")
        
        # Ollama Settings
        self.ollama_base_url: str = os.getenv("OLLAMA_BASE_URL", "http://localhost:11434")
        
        # Retry Settings
        self.max_retries: int = int(os.getenv("MAX_RETRIES", "3"))
        self.retry_min_wait: int = int(os.getenv("RETRY_MIN_WAIT", "2"))
        self.retry_max_wait: int = int(os.getenv("RETRY_MAX_WAIT", "10"))

        # MinIO Settings
        self.minio_endpoint: str = os.getenv("MINIO_ENDPOINT", "localhost:9000")
        self.minio_access_key: str = os.getenv("MINIO_ACCESS_KEY", "admin")
        self.minio_secret_key: str = os.getenv("MINIO_SECRET_KEY", "admin123")
        self.minio_bucket_name: str = os.getenv("MINIO_BUCKET_NAME", "tejomaya")
        self.minio_secure: bool = os.getenv("MINIO_SECURE", "False").lower() == "true"
        
        # Validate configuration
        self._validate()
    
    def _get_default_model(self) -> str:
        """Get default model based on provider."""
        provider = os.getenv("LLM_PROVIDER", "ollama").lower()
        if provider == "openai":
            return "gpt-4o-mini"  # Cost-effective GPT-4 variant
        elif provider == "azure":
            return "gpt-4o-mini"  # Azure deployment name
        elif provider == "aws":
            return "us.anthropic.claude-sonnet-4-5-20250929-v1:0"
        return "mistral"
    
    def _validate(self):
        """Validate configuration."""
        # If Azure is selected but no API key, fall back
        if self.llm_provider == "azure" and not self.azure_api_key:
            print("⚠️  Azure OpenAI selected but no API key found. Falling back to Ollama.")
            self.llm_provider = "ollama"
            self.llm_model = "mistral"
        # If OpenAI is selected but no API key, fall back to Ollama
        elif self.llm_provider == "openai" and not self.openai_api_key:
            print("⚠️  OpenAI selected but no API key found. Falling back to Ollama.")
            self.llm_provider = "ollama"
            self.llm_model = "mistral"
        # If AWS is selected, it implicitly relies on IAM roles or standard AWS env variables (e.g., AWS_PROFILE).
        
    def is_aws_available(self) -> bool:
        """Check if AWS Bedrock is configured and available (assumes AWS CLI/IAM is set up)."""
        return self.llm_provider == "aws"

    
    def is_openai_available(self) -> bool:
        """Check if OpenAI is configured and available."""
        return self.llm_provider == "openai" and bool(self.openai_api_key)
    
    def is_azure_available(self) -> bool:
        """Check if Azure OpenAI is configured and available."""
        return self.llm_provider == "azure" and bool(self.azure_api_key)

    def is_minio_configured(self) -> bool:
        """Check if MinIO is configured."""
        return bool(self.minio_endpoint) and bool(self.minio_access_key) and bool(self.minio_secret_key)
    
    def __repr__(self) -> str:
        """String representation (hides API key)."""
        return (
            f"Config(provider={self.llm_provider}, "
            f"model={self.llm_model}, "
            f"temperature={self.llm_temperature}, "
            f"openai_key={'***' if self.openai_api_key else 'not set'}, "
            f"azure_key={'***' if self.azure_api_key else 'not set'})"
        )


# Global configuration instance
_config = None


def get_config() -> Config:
    """Get or create global configuration instance."""
    global _config
    if _config is None:
        _config = Config()
    return _config


def reload_config() -> Config:
    """Reload configuration (useful for testing or config changes)."""
    global _config
    _config = Config()
    return _config
