"""
Interactive CLI for WMI Agent

Provides a command-line interface to interact with the WMI Agent.
Supports multiple providers: Ollama (local) and Azure AI Foundry (cloud).
"""

import asyncio
import sys
import os
from typing import Optional
from .agent import create_wmi_agent, ProviderType


class WMIAgentCLI:
    """Interactive CLI for WMI Agent"""
    
    def __init__(
        self,
        provider: Optional[ProviderType] = None,
        model_id: Optional[str] = None,
        endpoint: Optional[str] = None
    ):
        """
        Initialize the CLI
        
        Args:
            provider: Provider type - "ollama" or "azure" (default: from AGENT_PROVIDER env)
            model_id: Model identifier (uses env defaults if None)
            endpoint: API endpoint (uses env defaults if None)
        """
        self.provider = provider or os.getenv("AGENT_PROVIDER", "ollama")
        self.model_id = model_id
        self.endpoint = endpoint
        self.agent = None
    
    async def initialize(self):
        """Initialize the agent"""
        print("🤖 Initializing WMI Agent...")
        print(f"   Provider: {self.provider}")
        if self.model_id:
            print(f"   Model: {self.model_id}")
        if self.endpoint:
            print(f"   Endpoint: {self.endpoint}")
        print()
        
        try:
            self.agent = await create_wmi_agent(
                provider=self.provider,
                model_id=self.model_id,
                endpoint=self.endpoint
            )
            print("✓ Agent initialized successfully!\n")
            return True
        except Exception as e:
            print(f"✗ Error initializing agent: {e}")
            
            if self.provider == "ollama":
                print("\nMake sure Ollama is running and the model is installed:")
                print(f"  ollama pull {self.model_id or os.getenv('OLLAMA_MODEL', 'gpt-oss:120b')}")
                print(f"  ollama serve")
            elif self.provider == "azure":
                print("\nMake sure Azure OpenAI environment variables are set:")
                print("  AZURE_OPENAI_ENDPOINT: Your Azure OpenAI endpoint")
                print("  AZURE_OPENAI_API_KEY: Your Azure OpenAI API key")
                print("  AZURE_OPENAI_DEPLOYMENT: Deployment name (e.g., gpt-4.1)")
            
            return False
    
    def print_help(self):
        """Print help message"""
        print("\n" + "="*60)
        print("WMI Agent - Natural Language Windows Management")
        print("="*60)
        print("\nAvailable Commands:")
        print("  /help    - Show this help message")
        print("  /exit    - Exit the CLI")
        print("  /quit    - Exit the CLI")
        print("  /clear   - Clear the screen")
        print("\nExample Queries:")
        print("  - What's my current memory usage?")
        print("  - Show me running services")
        print("  - What's the CPU usage?")
        print("  - List disk drives and their space")
        print("  - Show network adapter configuration")
        print("  - What processes are using the most memory?")
        print("  - Get system uptime")
        print("  - Am I running as administrator?")
        print("\nTip: You can ask questions in natural language!")
        print("="*60 + "\n")
    
    async def run_interactive(self):
        """Run the interactive CLI"""
        # Initialize agent
        if not await self.initialize():
            return
        
        # Print welcome message
        self.print_help()
        
        try:
            while True:
                # Get user input
                try:
                    user_input = input("You: ").strip()
                except (EOFError, KeyboardInterrupt):
                    print("\n\nExiting...")
                    break
                
                if not user_input:
                    continue
                
                # Handle commands
                if user_input.startswith('/'):
                    command = user_input.lower()
                    
                    if command in ['/exit', '/quit']:
                        print("Goodbye!")
                        break
                    elif command == '/help':
                        self.print_help()
                        continue
                    elif command == '/clear':
                        import os
                        os.system('cls' if sys.platform == 'win32' else 'clear')
                        continue
                    else:
                        print(f"Unknown command: {user_input}")
                        print("Type /help for available commands")
                        continue
                
                # Process query with agent
                print("\n🤖 Agent: ", end='', flush=True)
                
                try:
                    response = await self.agent.run(user_input)
                    print(response)
                except Exception as e:
                    print(f"\n✗ Error: {e}")
                
                print()  # Empty line for readability
        
        finally:
            # Cleanup
            if self.agent:
                await self.agent.close()
                print("Agent closed.")
    
    async def run_single_query(self, query: str):
        """
        Run a single query and exit
        
        Args:
            query: The query to run
        """
        if not await self.initialize():
            return
        
        try:
            print(f"Query: {query}\n")
            print("Agent: ", end='', flush=True)
            
            response = await self.agent.run(query)
            print(response)
            print()
        
        except Exception as e:
            print(f"Error: {e}")
        
        finally:
            if self.agent:
                await self.agent.close()


async def main():
    """Main entry point"""
    import argparse
    
    parser = argparse.ArgumentParser(
        description="WMI Agent - Natural Language Windows Management",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Interactive mode (uses AGENT_PROVIDER env or defaults to Ollama)
  python cli_agent.py
  
  # Azure AI Inference
  python cli_agent.py --provider azure
  
  # Single query
  python cli_agent.py --query "What's the CPU usage?"
  
  # Different Ollama model
  python cli_agent.py --model llama3.1:8b
  
  # Custom Ollama endpoint
  python cli_agent.py --endpoint http://192.168.1.100:11434/v1

Environment Variables:
  AGENT_PROVIDER: Provider to use ("ollama" or "azure", default: ollama)
  
  For Ollama:
    OLLAMA_MODEL: Model name (default: gpt-oss:120b)
    OLLAMA_ENDPOINT: Ollama endpoint (default: http://localhost:11434/v1)
  
  For Azure OpenAI:
    AZURE_OPENAI_ENDPOINT: Azure OpenAI endpoint (required)
    AZURE_OPENAI_API_KEY: Azure OpenAI API key (required)
    AZURE_OPENAI_DEPLOYMENT: Deployment name (required, e.g., gpt-4.1)
    AZURE_OPENAI_API_VERSION: API version (default: 2024-08-01-preview)
"""
    )
    
    parser.add_argument(
        '--provider',
        choices=['ollama', 'azure'],
        help='Provider to use (default: from AGENT_PROVIDER env or "ollama")'
    )
    
    parser.add_argument(
        '--model',
        help='Model to use (default: from env based on provider)'
    )
    
    parser.add_argument(
        '--endpoint',
        help='API endpoint (default: from env based on provider)'
    )
    
    parser.add_argument(
        '--query',
        help='Single query to run (non-interactive mode)'
    )
    
    args = parser.parse_args()
    
    # Create CLI
    cli = WMIAgentCLI(
        provider=args.provider,
        model_id=args.model,
        endpoint=args.endpoint
    )
    
    # Run in appropriate mode
    if args.query:
        await cli.run_single_query(args.query)
    else:
        await cli.run_interactive()


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("\n\nInterrupted by user")
        sys.exit(0)


def cli_main():
    """Entry point for console script"""
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("\n\nInterrupted by user")
        sys.exit(0)
