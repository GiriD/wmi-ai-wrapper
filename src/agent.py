"""
WMI Agent using Microsoft Agent Framework with multiple providers

This module creates an AI agent that can interact with Windows Management Instrumentation (WMI)
using Microsoft Agent Framework with support for:
- Ollama (local LLM inference)
- Azure OpenAI (cloud-based models)
"""

import os
from typing import Optional, Literal
from dotenv import load_dotenv
from agent_framework import ChatAgent
from agent_framework.openai import OpenAIChatClient
from agent_framework.azure import AzureOpenAIChatClient
from .wmi_tools import get_wmi_tools

# Load environment variables from .env file
load_dotenv()

# Provider types
ProviderType = Literal["ollama", "azure"]


class WMIAgent:
    """
    AI Agent for Windows Management Instrumentation
    
    This agent uses Microsoft Agent Framework with support for:
    - Ollama: Local LLM inference
    - Azure AI Foundry: Cloud-based models (gpt-4o, etc.)
    """
    
    DEFAULT_INSTRUCTIONS = """You are a helpful Windows system administrator assistant with deep knowledge of WMI (Windows Management Instrumentation).

Your capabilities:
- Query and report Windows system information (OS, hardware, BIOS)
- Monitor system resources (CPU, memory, disk usage)
- Manage and query Windows services
- List and analyze running processes with CPU and memory metrics
- Retrieve network adapter configuration
- Execute custom WQL queries when needed

Guidelines:
- Always provide clear, concise answers
- When showing lists, limit to most relevant items
- Explain technical terms when appropriate
- For CPU usage by process, use get_process_performance tool which provides CPU percentages
- For memory-only process listings, use list_processes tool
- For administrative tasks, inform users if admin privileges are required
- Be proactive in suggesting related information that might be helpful
- Format output clearly with proper sections and bullet points

Remember: You have direct access to WMI through function tools, so always use them to get real, current data rather than making assumptions."""
    
    def __init__(
        self,
        provider: Optional[ProviderType] = None,
        model_id: Optional[str] = None,
        endpoint: Optional[str] = None,
        instructions: Optional[str] = None,
        name: str = "WMI Agent"
    ):
        """
        Initialize the WMI Agent
        
        Args:
            provider: Provider type - "ollama" or "azure" (default: from AGENT_PROVIDER env)
            model_id: Model identifier (default: from env based on provider)
            endpoint: API endpoint (default: from env based on provider)
            instructions: Custom agent instructions (uses default if None)
            name: Agent name
        
        Environment variables:
            AGENT_PROVIDER: Provider to use ("ollama" or "azure", default: "ollama")
            
            For Ollama:
                OLLAMA_MODEL: Model name (default: "gpt-oss:120b")
                OLLAMA_ENDPOINT: Ollama endpoint (default: "http://localhost:11434/v1")
            
            For Azure OpenAI:
                AZURE_OPENAI_ENDPOINT: Azure OpenAI endpoint (required)
                AZURE_OPENAI_API_KEY: Azure OpenAI API key (required)
                AZURE_OPENAI_DEPLOYMENT: Deployment name (required, e.g., gpt-4.1)
                AZURE_OPENAI_API_VERSION: API version (default: 2024-08-01-preview)
        """
        # Get provider from parameter or environment
        self.provider = provider or os.getenv("AGENT_PROVIDER", "ollama")
        if self.provider not in ["ollama", "azure"]:
            raise ValueError(f"Unsupported provider: {self.provider}. Use 'ollama' or 'azure'.")
        
        self.instructions = instructions or self.DEFAULT_INSTRUCTIONS
        self.name = name
        self._agent = None
        
        # Set defaults based on provider
        if self.provider == "ollama":
            self.model_id = model_id or os.getenv("OLLAMA_MODEL", "gpt-oss:120b")
            self.endpoint = endpoint or os.getenv("OLLAMA_ENDPOINT", "http://localhost:11434/v1")
        elif self.provider == "azure":
            # Azure OpenAI uses deployment name, not model_id
            self.deployment_name = model_id or os.getenv("AZURE_OPENAI_DEPLOYMENT")
            self.endpoint = endpoint or os.getenv("AZURE_OPENAI_ENDPOINT")
            self.api_version = os.getenv("AZURE_OPENAI_API_VERSION", "2024-08-01-preview")
            
            if not self.endpoint:
                raise ValueError(
                    "For Azure provider, AZURE_OPENAI_ENDPOINT environment variable is required"
                )
            if not self.deployment_name:
                raise ValueError(
                    "For Azure provider, AZURE_OPENAI_DEPLOYMENT environment variable is required"
                )
            
            # For Azure OpenAI, construct the full endpoint URL
            # Format: https://<resource>.openai.azure.com/openai/deployments/<deployment>/chat/completions?api-version=<version>
            if not self.endpoint.endswith('/'):
                self.endpoint += '/'
            self.model_id = self.deployment_name  # Agent framework uses model_id
    
    async def create_agent(self):
        """Create the agent instance based on provider"""
        # Get all WMI tools as standalone functions
        tools = get_wmi_tools()
        
        if self.provider == "ollama":
            # Create OpenAI-compatible client for Ollama
            # Note: Ollama doesn't require an API key, but OpenAIChatClient does
            # We provide a dummy key for local Ollama endpoints
            chat_client = OpenAIChatClient(
                model_id=self.model_id,
                base_url=self.endpoint,
                api_key="ollama"  # Dummy key for local Ollama (not validated)
            )
        
        elif self.provider == "azure":
            # Create Azure OpenAI client
            api_key = os.getenv("AZURE_OPENAI_API_KEY")
            if not api_key:
                raise ValueError(
                    "For Azure provider, AZURE_OPENAI_API_KEY environment variable is required"
                )
            
            # Use AzureOpenAIChatClient for Azure OpenAI
            chat_client = AzureOpenAIChatClient(
                ai_model_id=self.deployment_name,
                endpoint=self.endpoint,
                api_key=api_key,
                deployment_name=self.deployment_name
            )
        
        # Create the agent
        self._agent = ChatAgent(
            chat_client=chat_client,
            name=self.name,
            instructions=self.instructions,
            tools=tools
        )
        
        return self._agent
    
    def get_new_thread(self):
        """
        Create a new conversation thread for maintaining context
        
        Returns:
            AgentThread instance for conversation continuity
        """
        if not self._agent:
            raise RuntimeError("Agent not initialized. Call create_agent() first.")
        
        return self._agent.get_new_thread()
    
    async def run(self, message: str, thread=None) -> str:
        """
        Run a single query against the agent
        
        Args:
            message: User message/query
            thread: AgentThread instance for conversation context (optional)
        
        Returns:
            Agent's response text
        """
        if not self._agent:
            await self.create_agent()
        
        result = await self._agent.run(message, thread=thread)
        return result.text
    
    async def run_streaming(self, message: str, thread=None):
        """
        Run a query with streaming response
        
        Args:
            message: User message/query
            thread: AgentThread instance for conversation context (optional)
        
        Yields:
            Chunks of the agent's response
        """
        if not self._agent:
            await self.create_agent()
        
        async for chunk in self._agent.run_streaming(message, thread=thread):
            yield chunk
    
    async def close(self):
        """Close the agent and cleanup resources"""
        # ChatAgent doesn't have a close method, so we just clear the reference
        if self._agent:
            self._agent = None


async def create_wmi_agent(
    provider: ProviderType = "ollama",
    model_id: Optional[str] = None,
    endpoint: Optional[str] = None,
    instructions: Optional[str] = None
) -> WMIAgent:
    """
    Factory function to create and initialize a WMI Agent
    
    Args:
        provider: Provider type - "ollama" or "azure"
        model_id: Model identifier (uses defaults if None)
        endpoint: API endpoint (uses defaults if None)
        instructions: Custom agent instructions
    
    Returns:
        Initialized WMIAgent instance
    
    Examples:
        ```python
        # Ollama
        agent = await create_wmi_agent(provider="ollama", model_id="gpt-oss:120b")
        
        # Azure AI Foundry
        agent = await create_wmi_agent(provider="azure", model_id="gpt-4o")
        
        response = await agent.run("What's the current CPU usage?")
        print(response)
        await agent.close()
        ```
    """
    agent = WMIAgent(
        provider=provider,
        model_id=model_id,
        endpoint=endpoint,
        instructions=instructions
    )
    await agent.create_agent()
    return agent
