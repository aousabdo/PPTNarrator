import os
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

CHUNK_SIZE = int(os.getenv("CHUNK_SIZE", 1024))
XI_API_KEY = os.getenv("ELEVENLABS_API_KEY")
VOICE_ID = os.getenv("VOICE_ID")
MAX_WORKERS = int(os.getenv("MAX_WORKERS", 5))
ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY")

# Validate that required environment variables are set
if not XI_API_KEY:
    raise ValueError("ELEVENLABS_API_KEY is not set in the .env file")
if not ANTHROPIC_API_KEY:
    raise ValueError("ANTHROPIC_API_KEY is not set in the .env file")
if not VOICE_ID:
    raise ValueError("VOICE_ID is not set in the .env file")