"""
WMI Agent - AI-powered Windows Management Instrumentation

This package provides an AI agent for natural language interaction with Windows WMI.
"""

from .agent import WMIAgent, create_wmi_agent, ProviderType
from .wmi_tools import get_wmi_tools

__all__ = [
    "WMIAgent",
    "create_wmi_agent",
    "ProviderType",
    "get_wmi_tools",
]
