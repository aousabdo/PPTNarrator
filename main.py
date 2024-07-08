import argparse
import logging
import os
from concurrent.futures import ThreadPoolExecutor
from config import MAX_WORKERS
from text_to_speech import text_to_speech
from narration_generator import process_slides, add_audio_to_ppt

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def generate_narrations(ppt_path, output_dir):
    # Generate narrations and save as text files
    narrations = process_slides(ppt_path, output_dir)
    if not narrations:
        logging.error("No narrations were generated. Exiting.")
        return False
    
    logging.info(f"Narrations generated and saved in {output_dir}")
    return True

def generate_audio(output_dir):
    # Read edited narration files and generate audio
    narration_files = sorted([f for f in os.listdir(output_dir) if f.endswith('_narration.txt')])
    audio_files = []

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        for file in narration_files:
            with open(os.path.join(output_dir, file), 'r') as f:
                text = f.read()
            audio_path = os.path.join(output_dir, file.replace('_narration.txt', '.mp3'))
            executor.submit(text_to_speech, text, audio_path)
            audio_files.append(audio_path)

    logging.info(f"Audio files generated in {output_dir}")
    return audio_files

def insert_audio(ppt_path, audio_files):
    # Sort audio files to ensure correct order
    sorted_audio_files = sorted(audio_files, key=lambda x: int(os.path.basename(x).split('_')[1].split('.')[0]))
    # Add audio files to the PowerPoint
    final_ppt = add_audio_to_ppt(ppt_path, sorted_audio_files)
    logging.info(f"Final presentation with audio saved as: {final_ppt}")

def main():
    parser = argparse.ArgumentParser(description="Generate narrations for PowerPoint slides")
    parser.add_argument("ppt_path", help="Path to the PowerPoint file")
    parser.add_argument("output_dir", help="Directory to save the generated files")
    parser.add_argument("--generate-narrations", action="store_true", help="Generate narrations from slides")
    parser.add_argument("--generate-audio", action="store_true", help="Generate audio from narration files")
    parser.add_argument("--insert-audio", action="store_true", help="Insert audio into PowerPoint")
    args = parser.parse_args()

    os.makedirs(args.output_dir, exist_ok=True)

    if args.generate_narrations:
        if not generate_narrations(args.ppt_path, args.output_dir):
            return

    if args.generate_audio:
        audio_files = generate_audio(args.output_dir)
        if not audio_files:
            logging.error("No audio files were generated. Exiting.")
            return

    if args.insert_audio:
        audio_files = sorted([os.path.join(args.output_dir, f) for f in os.listdir(args.output_dir) if f.endswith('.mp3')],
                             key=lambda x: int(os.path.basename(x).split('_')[1].split('.')[0]))
        if not audio_files:
            logging.error("No audio files found. Exiting.")
            return
        insert_audio(args.ppt_path, audio_files)

if __name__ == "__main__":
    main()