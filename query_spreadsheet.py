#!/usr/bin/env python3
"""CLI interface for querying compressed spreadsheet data."""

import os
import argparse
from typing import Optional
from pathlib import Path

from Core.llm_query import LLMQuerySystem

def main():
    """Run the query interface."""
    parser = argparse.ArgumentParser(description="Query compressed spreadsheet data using LLM")
    parser.add_argument("--output-dir", default="output", help="Directory containing compressed files")
    parser.add_argument("--api-key", help="OpenAI API key (or set OPENAI_API_KEY env var)")
    parser.add_argument("--query", help="Query to process (optional)")
    args = parser.parse_args()
    
    # Get API key
    api_key = args.api_key or os.getenv("OPENAI_API_KEY")
    if not api_key:
        print("Error: OpenAI API key not provided. Set it via --api-key or OPENAI_API_KEY env var.")
        return
    
    # Initialize query system
    try:
        query_system = LLMQuerySystem(args.output_dir, api_key)
    except Exception as e:
        print(f"Error initializing query system: {e}")
        return
    
    # Process query if provided
    if args.query:
        response = query_system.process_query(args.query)
        print("\nResponse:")
        print(response)
        return
    
    # Interactive mode
    print("Enter your queries (type 'exit' to quit):")
    while True:
        try:
            query = input("\nQuery: ").strip()
            if query.lower() in ['exit', 'quit']:
                break
            
            if not query:
                continue
            
            response = query_system.process_query(query)
            print("\nResponse:")
            print(response)
            
        except KeyboardInterrupt:
            print("\nExiting...")
            break
        except Exception as e:
            print(f"Error: {e}")

if __name__ == "__main__":
    main() 