import os
import logging
import base64
import anthropic

ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY")

def get_narration_from_claude(image_path, slide_number, total_slides, presentation_summary):
    if not ANTHROPIC_API_KEY:
        raise ValueError("ANTHROPIC_API_KEY environment variable is not set")

    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    with open(image_path, "rb") as image_file:
        base64_image = base64.b64encode(image_file.read()).decode('utf-8')

    prompt = f"""You are an AI assistant tasked with creating speaker notes for slide {slide_number} of a {total_slides}-slide presentation, optimized for narration using ElevenLabs AI text-to-speech technology. Your goal is to transform the given slide content into engaging, conversational speaker notes that flow naturally when read aloud.

Here's a summary of the entire presentation to provide context:

{presentation_summary}

Now, focus on the current slide. Your goal is to transform the given slide content into engaging, conversational speaker notes that flow naturally when read aloud, keeping in mind the overall context of the presentation.

Follow these guidelines to create effective speaker notes:

1. Aim for a slide narration time of about two minutes, unless if the slide is the first slide, then aim for 1 minute or less.
2. Use a conversational tone that sounds natural when spoken.
3. Vary your opening phrases for each slide. Avoid starting every slide with phrases like "Now, let's..." or "Moving on to...". Instead, use a variety of transitions that feel natural and fit the content.
4. Focus on clear explanations of the content without mentioning "Slide X" or "Text Block" prefixes.
5. Provide detailed, but not too detailed, descriptions of any figures, charts, or images on the slide.
6. For code snippets, explain the purpose and functionality of the code in a way that's easy to understand when spoken.
7. Vary the pacing to emphasize key points.
8. Incorporate natural pauses using the <break time="Xs" /> syntax, where X is the duration in seconds. For example:
   - Use short pauses (0.2-0.5 seconds) between sentences or phrases.
   - Use medium pauses (0.5-1 second) between main ideas.
   - Use longer pauses (1-2 seconds) for transitions between major topics.
9. To put emphasis on a word or a phrase, use double quotes and caps. For example:
   - The universe is 13.8 "BILLION" years old
10. Ensure the narration is engaging and informative, suitable for an educational context.
11. Expand on the slide content where necessary to provide context or additional information.
12. Use transitional phrases to connect ideas and maintain flow.
13. Do not include any prefixes like "Slide X", "Text Block", or mention that you're describing a specific slide.
14. If this is not the first slide, do not start with "Welcome" or act as if this is the beginning of the presentation. Instead, use appropriate transitions based on the slide number.
15. If this is the last slide, provide a suitable conclusion to the presentation.

Here is the slide content to work with:

<slide_content>
[An image of the slide content is attached]
</slide_content>

Your task is to provide optimized speaker notes for this slide. Include appropriate breaks using the <break time="Xs" /> syntax throughout your notes. Remember to use a conversational tone, explain the content clearly, and incorporate pauses and emphasis as needed to create engaging and natural-sounding narration that flows well with the rest of the presentation."""

    try:
        message = client.messages.create(
            model="claude-3-5-sonnet-20240620",
            max_tokens=500,
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
        
        # Extract the text content from the ContentBlock
        if message.content and len(message.content) > 0:
            content_block = message.content[0]
            if hasattr(content_block, 'text'):
                return content_block.text
            else:
                return str(content_block)
        else:
            return "No content generated."
    except Exception as e:
        logging.error(f"Error calling Anthropic API: {str(e)}")
        return None

def get_summary_from_claude(full_text):
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    prompt = f"""You are an AI assistant tasked with summarizing a presentation. Here's the full text of the presentation:

{full_text}

Please provide a concise summary of the main points and overall structure of this presentation. This summary will be used to provide context for narrating individual slides."""

    try:
        message = client.messages.create(
            model="claude-3-5-sonnet-20240620",
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
    image_path = "path/to/your/slide/image.png"
    narration = get_narration_from_claude(image_path)
    if narration:
        print(narration)
    else:
        print("Failed to generate narration.")