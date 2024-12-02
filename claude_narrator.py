import os
import logging
import base64
import json
import re
import anthropic

ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY")

def get_narrations_from_claude(image_paths, presentation_summary):
    if not ANTHROPIC_API_KEY:
        raise ValueError("ANTHROPIC_API_KEY environment variable is not set")

    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    all_narrations = {}
    used_openings = set()

    for i, image_path in enumerate(image_paths, 1):
        with open(image_path, "rb") as image_file:
            base64_image = base64.b64encode(image_file.read()).decode('utf-8')

        prompt = f"""You are an AI assistant tasked with creating speaker notes for slide {i} of a {len(image_paths)}-slide presentation on Exploratory Data Analysis in Data Science. Your goal is to transform the given slide content into engaging, conversational speaker notes that flow naturally when read aloud, as if you are the presenter speaking directly to the audience.

Here's a summary of the entire presentation to provide overall context:
{presentation_summary}

Here is the content for the current slide (slide {i}):

[An image of the slide content is attached]

Your task is to provide optimized speaker notes for this specific slide, written from the perspective of the presenter. The narration should directly correspond to the content visible in the image.

Follow these guidelines:

1. The narration should be in the first person, using "I" and "we" naturally.
2. Do not mention or introduce any names that appear on the slides.
3. Aim for about one to two minutes of narration for this slide.
4. Focus on clear, concise explanations of the specific content shown on the slide.
5. Describe any visible figures, charts, or images on the slide without redundancy.
6. Use the <break time="Xs" /> syntax for pauses, where X is the duration in seconds.
7. Use double quotes and caps for emphasis. Example: This finding is "CRUCIAL" for our understanding.
8. Ensure the narration is engaging and informative, suitable for an educational context.
9. Create a unique opening for this slide. Do not start with "Welcome", "Now, let's", or any phrase you've used for previous slides.
10. If this is the last slide, provide a suitable conclusion that ties back to the overall presentation themes.
11. Be concise and avoid unnecessary repetition. Each sentence should provide new information or insight.
12. If you must repeat a point for emphasis, do so intentionally and sparingly, using different phrasing.
13. Minimize the use of filler phrases like "Alright," "Let's dive into," etc. Start sentences with varied, engaging openings.
14. Ensure that your narration follows a logical flow, with each point building on the previous one.
15. Tailor your language to the complexity of the content. Use simpler explanations for basic concepts and more technical language for advanced topics.

Format your response as follows:

[START_NARRATION]
(Your narration here, directly related to the slide's specific content)
[END_NARRATION]
"""

        try:
            message = client.messages.create(
                model="claude-3-5-sonnet-latest",
                max_tokens=1000,
                temperature=0,
                messages=[
                    {
                        "role": "user",
                        "content": [
                            {
                                "type": "image",
                                "source": {
                                    "type": "base64",
                                    "media_type": "image/png",
                                    "data": base64_image
                                }
                            },
                            {
                                "type": "text",
                                "text": prompt
                            }
                        ]
                    }
                ]
            )
            
            if message.content and len(message.content) > 0:
                content = message.content[0].text
                start_tag = '[START_NARRATION]'
                end_tag = '[END_NARRATION]'
                if start_tag in content and end_tag in content:
                    narration = content.split(start_tag)[1].split(end_tag)[0].strip()
                    
                    # Check for repetitive openings
                    first_sentence = narration.split('.')[0]
                    if any(first_sentence.startswith(opening) for opening in used_openings):
                        narration = "For this slide, " + narration
                    else:
                        used_openings.add(first_sentence)
                    
                    all_narrations[f"slide_{i}"] = {"narration": narration}
                else:
                    logging.warning(f"No properly formatted narration generated for slide {i}")
            else:
                logging.warning(f"No content generated for slide {i}")
        except Exception as e:
            logging.error(f"Error calling Anthropic API for slide {i}: {str(e)}")

    return all_narrations

def get_summary_from_claude(full_text):
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    prompt = f"""You are an AI assistant tasked with summarizing a presentation. Here's the full text of the presentation:

    {full_text}

    Please provide a concise summary of the main points and overall structure of this presentation. This summary will be used to provide context for narrating individual slides."""

    try:
        message = client.messages.create(
            model="claude-3-5-sonnet-latest",
            max_tokens=500,
            temperature=0,
            messages=[
                {
                    "role": "user",
                    "content": prompt
                }
            ]
        )
        
        if message.content and len(message.content) > 0:
            return message.content[0].text
        else:
            return "Unable to generate summary."
    except Exception as e:
        logging.error(f"Error calling Anthropic API for summary: {str(e)}")
        return None

# Example usage
if __name__ == "__main__":
    image_paths = ["path/to/slide1.png", "path/to/slide2.png", ...]  # Add all slide image paths
    presentation_summary = "This presentation covers..."  # Add your presentation summary
    narrations = get_narrations_from_claude(image_paths, presentation_summary)
    if narrations:
        print(json.dumps(narrations, indent=2))
    else:
        print("Failed to generate narrations.")