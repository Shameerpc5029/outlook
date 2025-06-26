#!/usr/bin/env python3
"""
Outlook MCP Server - Main entry point

This file provides the main entry point for the Outlook MCP server.
It can be run directly or used as the entry point for the package.
"""

import sys
import os

# Add the src directory to the path so we can import our modules
current_dir = os.path.dirname(os.path.abspath(__file__))
src_dir = os.path.join(current_dir, 'src')
if src_dir not in sys.path:
    sys.path.insert(0, src_dir)

def main():
    """Main entry point for the Outlook MCP server."""
    from outlook_mcp.server import main as server_main
    server_main()

if __name__ == "__main__":
    main()
