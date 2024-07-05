import requests
import logging
from config import CHUNK_SIZE, XI_API_KEY, VOICE_ID

def text_to_speech(text_to_speak, output_path):
    tts_url = f"https://api.elevenlabs.io/v1/text-to-speech/{VOICE_ID}/stream"
    
    headers = {
        "Accept": "application/json",
        "xi-api-key": XI_API_KEY
    }
    
    data = {
        "text": text_to_speak,
        "model_id": "eleven_multilingual_v2",
        "voice_settings": {
            "stability": 0.5,
            "similarity_boost": 0.8,
            "style": 0.0,
            "use_speaker_boost": True
        }
    }
    
    try:
        response = requests.post(tts_url, headers=headers, json=data, stream=True)
        response.raise_for_status()
        
        with open(output_path, "wb") as f:
            for chunk in response.iter_content(chunk_size=CHUNK_SIZE):
                f.write(chunk)
        logging.info(f"Audio stream saved successfully to {output_path}")
    except requests.RequestException as e:
        logging.error(f"Error in text_to_speech for {output_path}: {str(e)}")