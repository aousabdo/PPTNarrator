import argparse
import logging
import os
from concurrent.futures import ThreadPoolExecutor
from config import MAX_WORKERS
from text_to_speech import text_to_speech
from narration_generator import process_slides, add_audio_to_ppt

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def main(ppt_path, output_dir):
    # Ensure output directory exists
    os.makedirs(output_dir, exist_ok=True)
    
    # Generate narrations and output paths
    narrations = process_slides(ppt_path, output_dir)
    
    if not narrations:
        logging.error("No narrations were generated. Exiting.")
        return
    
    # Generate audio for each narration using multiple threads
    audio_files = []
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = [executor.submit(text_to_speech, text, path) for text, path in narrations]
        for future in futures:
            future.result()  # This will raise any exceptions that occurred during execution
    
    # Collect paths of generated audio files
    audio_files = [path for _, path in narrations]
    
    # Add audio files to the PowerPoint
    final_ppt = add_audio_to_ppt(ppt_path, audio_files)
    logging.info(f"Final presentation with audio saved as: {final_ppt}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Generate audio narrations for PowerPoint slides and add them to the presentation")
    parser.add_argument("ppt_path", help="Path to the PowerPoint file")
    parser.add_argument("output_dir", help="Directory to save the generated audio files")
    args = parser.parse_args()

    main(args.ppt_path, args.output_dir)