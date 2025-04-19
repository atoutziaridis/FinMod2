"""LLM query system for compressed spreadsheet data."""

import os
import json
import logging
from typing import Dict, List, Optional, Any, Tuple
from pathlib import Path
from openai import OpenAI
from datetime import datetime
import hashlib

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

class MetadataIndex:
    """Manages metadata for compressed spreadsheet files."""
    
    def __init__(self, output_dir: str):
        """Initialize metadata index.
        
        Args:
            output_dir: Directory containing compressed files
        """
        self.output_dir = output_dir
        self.metadata_file = os.path.join(output_dir, "metadata_index.json")
        self.index = self._load_or_create_index()
    
    def _load_or_create_index(self) -> Dict:
        """Load existing index or create new one."""
        if os.path.exists(self.metadata_file):
            try:
                with open(self.metadata_file, 'r') as f:
                    return json.load(f)
            except Exception as e:
                logger.error(f"Error loading metadata index: {e}")
                return {}
        return {}
    
    def update_index(self, sheet_name: str, metadata: Dict) -> None:
        """Update index with new sheet metadata.
        
        Args:
            sheet_name: Name of the sheet
            metadata: Metadata about the sheet
        """
        self.index[sheet_name] = {
            "last_updated": datetime.now().isoformat(),
            **metadata
        }
        self._save_index()
    
    def _save_index(self) -> None:
        """Save current index to file."""
        try:
            with open(self.metadata_file, 'w') as f:
                json.dump(self.index, f, indent=2)
        except Exception as e:
            logger.error(f"Error saving metadata index: {e}")
    
    def get_relevant_files(self, query: str) -> List[str]:
        """Get list of relevant files based on query.
        
        Args:
            query: User query string
            
        Returns:
            List of relevant file paths
        """
        relevant_files = []
        for sheet_name, metadata in self.index.items():
            if self._is_relevant(metadata, query):
                file_path = os.path.join(self.output_dir, f"{sheet_name}_compressed.json")
                if os.path.exists(file_path):
                    relevant_files.append(file_path)
        return relevant_files
    
    def _is_relevant(self, metadata: Dict, query: str) -> bool:
        """Check if metadata is relevant to query.
        
        Args:
            metadata: Sheet metadata
            query: User query
            
        Returns:
            True if relevant, False otherwise
        """
        query = query.lower()
        relevance_score = 0
        
        # Check headers
        for header in metadata.get("headers", []):
            if header.lower() in query:
                relevance_score += 2
        
        # Check key metrics
        for metric in metadata.get("key_metrics", []):
            if metric.lower() in query:
                relevance_score += 3
        
        # Check date ranges if query mentions time periods
        if any(time_word in query for time_word in ["year", "month", "quarter", "date"]):
            if metadata.get("date_ranges"):
                relevance_score += 1
        
        return relevance_score >= 2

class LLMQuerySystem:
    """System for querying compressed spreadsheet data using LLM."""
    
    def __init__(self, output_dir: str, api_key: str):
        """Initialize query system.
        
        Args:
            output_dir: Directory containing compressed files
            api_key: OpenAI API key
        """
        self.output_dir = output_dir
        self.metadata_index = MetadataIndex(output_dir)
        self.client = OpenAI(api_key=api_key)
        self.prompt_cache = {}
    
    def process_query(self, query: str) -> str:
        """Process user query and return response.
        
        Args:
            query: User query string
            
        Returns:
            Response string
        """
        try:
            # Get relevant files
            relevant_files = self.metadata_index.get_relevant_files(query)
            if not relevant_files:
                return "No relevant data found for your query."
            
            # Load relevant data
            relevant_data = []
            for file_path in relevant_files:
                try:
                    with open(file_path, 'r') as f:
                        data = json.load(f)
                        relevant_data.append(data)
                except Exception as e:
                    logger.error(f"Error loading file {file_path}: {e}")
            
            # Prepare prompt and get response
            prompt, cache_key = self._prepare_prompt(query, relevant_data)
            response = self._get_llm_response(prompt, cache_key)
            
            return response
            
        except Exception as e:
            logger.error(f"Error processing query: {e}")
            return "Sorry, I encountered an error processing your query."
    
    def _prepare_prompt(self, query: str, data: List[Dict]) -> Tuple[str, str]:
        """Prepare prompt for LLM and generate cache key.
        
        Args:
            query: User query
            data: Relevant data
            
        Returns:
            Tuple of (prompt string, cache key)
        """
        # Create a stable cache key based on query and data structure
        cache_key = hashlib.md5(
            (query + json.dumps(data, sort_keys=True)).encode()
        ).hexdigest()
        
        # Check cache
        if cache_key in self.prompt_cache:
            return self.prompt_cache[cache_key], cache_key
        
        # Prepare context
        context = "financial spreadsheet data containing: "
        if data:
            first_sheet = data[0]
            if "metadata" in first_sheet:
                metadata = first_sheet["metadata"]
                if "headers" in metadata:
                    context += f"columns like {', '.join(metadata['headers'][:5])}"
                if "date_ranges" in metadata:
                    context += f", covering dates from {metadata['date_ranges'].get('min')} to {metadata['date_ranges'].get('max')}"
        
        prompt = f"""You are a financial data assistant analyzing {context}.

Question: {query}

Data:
{json.dumps(data, indent=2)}

Please provide a clear and concise answer based on the data. If the data doesn't contain the information needed to answer the question, say so."""
        
        # Cache the prompt
        self.prompt_cache[cache_key] = prompt
        return prompt, cache_key
    
    def _get_llm_response(self, prompt: str, cache_key: str) -> str:
        """Get response from LLM with optimized configuration.
        
        Args:
            prompt: Formatted prompt
            cache_key: Cache key for the prompt
            
        Returns:
            Response string
        """
        try:
            response = self.client.chat.completions.create(
                model="gpt-4.1-mini",
                messages=[
                    {"role": "system", "content": "You are a very technical and experiencedfinancial data assistant."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.3,  # Lower temperature for more consistent responses
                max_tokens=30000,  # Increased token limit for complex queries and detailed analysis
                headers={'X-Cache-Hint': f'financial_query:{cache_key}'}  # Cache hint for cost optimization
            )
            return response.choices[0].message.content
        except Exception as e:
            logger.error(f"Error getting LLM response: {e}")
            return "Sorry, I encountered an error getting a response." 