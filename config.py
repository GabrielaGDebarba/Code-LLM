
import random
from keys import *
import requests

# Mode
mode = "cloudflare"  # Only cloudflare mode supported now

class CloudflareAI:
    def __init__(self, account_id, api_key, model):
        self.account_id = account_id
        self.api_key = api_key
        self.model = model
        # Correct URL format with model in path
        self.base_url = f"https://api.cloudflare.com/client/v4/accounts/{account_id}/ai/run/{model}"
        
    def run(self, messages, max_tokens=500, **kwargs):
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json"
        }
        data = {
            "messages": messages,
            "max_tokens": max_tokens
        }
        
        # Allow overriding params via kwargs
        data.update(kwargs)
        
        response = requests.post(self.base_url, headers=headers, json=data)
        response.raise_for_status()
        return response.json()

# Initialize Cloudflare client
cloudflare_model = "@cf/meta/llama-3.3-70b-instruct-fp8-fast"  # Using a more widely available model
client = CloudflareAI(CLOUDFLARE_ACCOUNT_ID, CLOUDFLARE_API_KEY, cloudflare_model)

# Embedding model for reference (if needed later)
embedding_model = "@cf/baai/bge-base-en-v1.5"