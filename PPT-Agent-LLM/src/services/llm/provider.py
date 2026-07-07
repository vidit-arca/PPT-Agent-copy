"""
Unified LLM provider interface with support for OpenAI and Ollama.
Includes retry logic, structured outputs, and automatic fallback.
"""

import json
import logging
from typing import Optional, Union
from tenacity import retry, stop_after_attempt, wait_exponential, retry_if_exception_type

from langchain_core.messages import SystemMessage, HumanMessage
from langchain_ollama import ChatOllama

from src.config.settings import get_config
from src.models.domain import MappingResponse
from src.services.llm.prompts import get_mapping_system_prompt, get_mapping_user_prompt, get_json_schema


class LLMProvider:
    """Base class for LLM providers."""
    
    def __init__(self, model: str, temperature: float = 0):
        self.model = model
        self.temperature = temperature
        self.logger = logging.getLogger(__name__)
    
    def invoke(self, messages: list) -> str:
        """Invoke the LLM with messages. Returns response content."""
        raise NotImplementedError
    
    def invoke_structured(self, messages: list, response_model: type) -> dict:
        """Invoke with structured output. Returns validated dict."""
        raise NotImplementedError


class OpenAIProvider(LLMProvider):
    """OpenAI GPT-4 provider with structured outputs."""
    
    def __init__(self, model: str, temperature: float = 0, api_key: str = None, base_url: str = None):
        super().__init__(model, temperature)
        
        try:
            from langchain_openai import ChatOpenAI
            
            self.client = ChatOpenAI(
                model=model,
                temperature=temperature,
                api_key=api_key,
                base_url=base_url
            )
            self.logger.info(f"✓ Initialized OpenAI provider with model: {model}")
        except ImportError:
            raise ImportError(
                "OpenAI provider requires langchain-openai. "
                "Install with: pip install langchain-openai"
            )
        except Exception as e:
            raise RuntimeError(f"Failed to initialize OpenAI provider: {e}")
    
    def invoke(self, messages: list) -> str:
        """Invoke OpenAI with messages."""
        response = self.client.invoke(messages)
        return response.content
    
    def invoke_structured(self, messages: list, response_model: type) -> dict:
        """
        Invoke with structured output using OpenAI's JSON mode.
        Returns validated Pydantic model as dict.
        """
        try:
            # Use OpenAI's structured output feature
            structured_llm = self.client.with_structured_output(response_model)
            result = structured_llm.invoke(messages)
            
            # Validate and convert to dict
            if isinstance(result, response_model):
                return result.to_dict() if hasattr(result, 'to_dict') else result.model_dump()
            
            return result
            
        except Exception as e:
            self.logger.error(f"Structured output failed: {e}")
            # Fallback to regular invoke and manual parsing
            response_text = self.invoke(messages)
            return self._parse_json_response(response_text, response_model)
    
    def _parse_json_response(self, text: str, response_model: type) -> dict:
        """Parse and validate JSON response."""
        import re
        
        # Try to extract JSON from markdown code blocks
        json_match = re.search(r'```(?:json)?\s*(\{.*?\})\s*```', text, re.DOTALL)
        if json_match:
            text = json_match.group(1)
        
        # Try to find JSON object
        json_match = re.search(r'\{.*\}', text, re.DOTALL)
        if json_match:
            text = json_match.group(0)
        
        # Parse and validate
        data = json.loads(text)
        validated = response_model(**data)
        return validated.to_dict() if hasattr(validated, 'to_dict') else validated.model_dump()


class AzureOpenAIProvider(LLMProvider):
    """Azure OpenAI provider with structured outputs."""
    
    def __init__(self, deployment: str, temperature: float = 0, api_key: str = None, 
                 azure_endpoint: str = None, api_version: str = None):
        super().__init__(deployment, temperature)
        
        try:
            from langchain_openai import AzureChatOpenAI
            
            self.client = AzureChatOpenAI(
                azure_deployment=deployment,
                api_version=api_version,
                azure_endpoint=azure_endpoint,
                api_key=api_key,
                temperature=temperature
            )
            self.logger.info(f"✓ Initialized Azure OpenAI provider with deployment: {deployment}")
        except ImportError:
            raise ImportError(
                "Azure OpenAI provider requires langchain-openai. "
                "Install with: pip install langchain-openai"
            )
        except Exception as e:
            raise RuntimeError(f"Failed to initialize Azure OpenAI provider: {e}")
    
    def invoke(self, messages: list) -> str:
        """Invoke Azure OpenAI with messages."""
        response = self.client.invoke(messages)
        return response.content
    
    def invoke_structured(self, messages: list, response_model: type) -> dict:
        """
        Invoke with structured output using Azure OpenAI's JSON mode.
        Returns validated Pydantic model as dict.
        """
        try:
            # Use Azure OpenAI's structured output feature
            structured_llm = self.client.with_structured_output(response_model)
            result = structured_llm.invoke(messages)
            
            # Validate and convert to dict
            if isinstance(result, response_model):
                return result.to_dict() if hasattr(result, 'to_dict') else result.model_dump()
            
            return result
            
        except Exception as e:
            self.logger.error(f"Structured output failed: {e}")
            # Fallback to regular invoke and manual parsing
            response_text = self.invoke(messages)
            return self._parse_json_response(response_text, response_model)
    
    def _parse_json_response(self, text: str, response_model: type) -> dict:
        """Parse and validate JSON response."""
        import re
        
        # Try to extract JSON from markdown code blocks
        json_match = re.search(r'```(?:json)?\s*(\{.*?\})\s*```', text, re.DOTALL)
        if json_match:
            text = json_match.group(1)
        
        # Try to find JSON object
        json_match = re.search(r'\{.*\}', text, re.DOTALL)
        if json_match:
            text = json_match.group(0)
        
        # Parse and validate
        data = json.loads(text)
        validated = response_model(**data)
        return validated.to_dict() if hasattr(validated, 'to_dict') else validated.model_dump()


class AWSBedrockProvider(LLMProvider):
    """AWS Bedrock provider with structured outputs."""
    
    def __init__(self, model_id: str, temperature: float = 0):
        super().__init__(model_id, temperature)
        
        try:
            from langchain_aws import ChatBedrock
            
            # Using us-east-1 since it's the standard for cross-region inference profiles
            self.client = ChatBedrock(
                model_id=model_id,
                model_kwargs={"temperature": temperature},
                region_name="us-east-1"
            )
            self.logger.info(f"✓ Initialized AWS Bedrock provider with model: {model_id}")
        except ImportError:
            raise ImportError(
                "AWS Bedrock provider requires langchain-aws. "
                "Install with: pip install langchain-aws boto3"
            )
        except Exception as e:
            raise RuntimeError(f"Failed to initialize AWS Bedrock provider: {e}")
    
    def invoke(self, messages: list) -> str:
        """Invoke AWS Bedrock with messages."""
        response = self.client.invoke(messages)
        return response.content
    
    def invoke_structured(self, messages: list, response_model: type) -> dict:
        """
        Invoke with structured output using Bedrock's structured output capability.
        Returns validated Pydantic model as dict.
        """
        try:
            # Claude 3.5 Sonnet supports tool use / structured outputs
            structured_llm = self.client.with_structured_output(response_model)
            result = structured_llm.invoke(messages)
            
            if isinstance(result, response_model):
                return result.to_dict() if hasattr(result, 'to_dict') else result.model_dump()
            
            return result
            
        except Exception as e:
            self.logger.error(f"Structured output failed: {e}")
            # Fallback to regular invoke and manual parsing
            response_text = self.invoke(messages)
            import re
            json_match = re.search(r'```(?:json)?\s*(\{.*?\})\s*```', response_text, re.DOTALL)
            if json_match:
                response_text = json_match.group(1)
            json_match = re.search(r'\{.*\}', response_text, re.DOTALL)
            if json_match:
                response_text = json_match.group(0)
            import json
            data = json.loads(response_text)
            validated = response_model(**data)
            return validated.to_dict() if hasattr(validated, 'to_dict') else validated.model_dump()


class OllamaProvider(LLMProvider):
    """Ollama provider for local models (Mistral, Llama, etc)."""
    
    def __init__(self, model: str, temperature: float = 0, base_url: str = "http://localhost:11434"):
        super().__init__(model, temperature)
        
        try:
            self.client = ChatOllama(
                model=model,
                temperature=temperature,
                base_url=base_url
            )
            # Test connection
            self.client.invoke([HumanMessage(content="ping")])
            self.logger.info(f"✓ Initialized Ollama provider with model: {model}")
        except Exception as e:
            raise RuntimeError(f"Failed to initialize Ollama provider: {e}")
    
    def invoke(self, messages: list) -> str:
        """Invoke Ollama with messages."""
        response = self.client.invoke(messages)
        return response.content
    
    def invoke_structured(self, messages: list, response_model: type) -> dict:
        """
        Invoke with structured output (manual parsing for Ollama).
        """
        response_text = self.invoke(messages)
        return self._parse_json_response(response_text, response_model)
    
    def _parse_json_response(self, text: str, response_model: type) -> dict:
        """Parse and validate JSON response."""
        import re
        
        # Clean up response
        text = text.strip()
        
        # Try to extract JSON from markdown code blocks
        json_match = re.search(r'```(?:json)?\s*(\{.*?\})\s*```', text, re.DOTALL)
        if json_match:
            text = json_match.group(1)
        
        # Try to find JSON object
        json_match = re.search(r'\{.*\}', text, re.DOTALL)
        if json_match:
            text = json_match.group(0)
        
        # Parse and validate
        data = json.loads(text)
        validated = response_model(**data)
        return validated.to_dict() if hasattr(validated, 'to_dict') else validated.model_dump()


def get_llm_provider() -> LLMProvider:
    """
    Factory function to get configured LLM provider.
    Automatically selects between Azure OpenAI, OpenAI, and Ollama based on config.
    """
    config = get_config()
    logger = logging.getLogger(__name__)
    
    # Try AWS Bedrock first if configured
    if config.is_aws_available():
        try:
            return AWSBedrockProvider(
                model_id=config.llm_model,
                temperature=config.llm_temperature
            )
        except Exception as e:
            logger.warning(f"⚠️  Failed to initialize AWS Bedrock provider: {e}")
            logger.info("Falling back to other providers...")
    

    
    # Try Azure OpenAI first if configured
    if config.is_azure_available():
        try:
            return AzureOpenAIProvider(
                deployment=config.azure_deployment,
                temperature=config.llm_temperature,
                api_key=config.azure_api_key,
                azure_endpoint=config.azure_endpoint,
                api_version=config.azure_api_version
            )
        except Exception as e:
            logger.warning(f"⚠️  Failed to initialize Azure OpenAI provider: {e}")
            logger.info("Falling back to Ollama provider...")
    
    # Try OpenAI if configured
    if config.is_openai_available():
        try:
            return OpenAIProvider(
                model=config.llm_model,
                temperature=config.llm_temperature,
                api_key=config.openai_api_key,
                base_url=config.openai_base_url
            )
        except Exception as e:
            logger.warning(f"⚠️  Failed to initialize OpenAI provider: {e}")
            logger.info("Falling back to Ollama provider...")
    
    # Fallback to Ollama
    try:
        return OllamaProvider(
            model=config.llm_model if config.llm_provider == "ollama" else "mistral",
            temperature=config.llm_temperature,
            base_url=config.ollama_base_url
        )
    except Exception as e:
        logger.error(f"✗ Failed to initialize Ollama provider: {e}")
        raise RuntimeError("No LLM provider available. Please check your configuration.")


@retry(
    stop=stop_after_attempt(3),
    wait=wait_exponential(multiplier=1, min=2, max=10),
    retry=retry_if_exception_type((json.JSONDecodeError, ValueError, KeyError))
)
def ask_llm_for_mapping_with_retry(
    provider: LLMProvider,
    template_spec: dict,
    data_profile: dict
) -> dict:
    """
    Ask LLM for mapping with automatic retry on failures.
    Uses structured outputs for better reliability.
    """
    logger = logging.getLogger(__name__)
    
    # Prepare messages
    system_prompt = get_mapping_system_prompt()
    user_prompt = get_mapping_user_prompt(template_spec, data_profile)
    
    messages = [
        SystemMessage(content=system_prompt),
        HumanMessage(content=user_prompt)
    ]
    
    # Try structured output first
    try:
        logger.debug("Attempting structured output...")
        result = provider.invoke_structured(messages, MappingResponse)
        logger.info("✓ Structured output validation: SUCCESS")
        return result
    except Exception as e:
        logger.warning(f"Structured output failed: {e}, trying regular invoke...")
        
        # Fallback to regular invoke
        response_text = provider.invoke(messages)
        
        # Manual JSON extraction and validation
        import re
        json_match = re.search(r'\{.*\}', response_text, re.DOTALL)
        if json_match:
            response_text = json_match.group(0)
        
        data = json.loads(response_text)
        validated = MappingResponse(**data)
        return validated.to_dict()
