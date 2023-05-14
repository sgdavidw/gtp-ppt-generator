import openai
import os
import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import time

from dotenv import load_dotenv

# Load .env file
load_dotenv()

# Retrieve API key from environment variables
openai.api_key = os.getenv("OPENAI_API_KEY")

template_path = 'templates/template.pptx'

# the folder to store the pptx file:
output_directory = 'ppt_files'

# powerpoint launguage
launguage = "English"

slide_notes = []

# if step_by_step is True, the script will generate a powerpoint file only with the outline but without the notes first; the script will not call openai API to generate the notes until you press 'C' to continue.
# if step_by_step is Fale, the script will generate the powerpoint file with the notes.
step_by_step = False

###put your presenation subject and background information here: 
ppt_title = "NIST Cybersecurity Framework"
additional_info = "1. what is NIST Cybersecurity Framework; 2. how it works; 3. how NIST Cybersecurity Framework can enterprises to improve security. 4.what the the differences between NIST Cybersecurity Framework and other cybersecurity framework like MITRE ATT&CK Framework. "
###

def generate_pptx(outlines, ppt_title, template_path, output_directory, add_notes=False):
    # Create a new PowerPoint presentation from template
    # and fill it with generated key points (slides) and text.
    # See 'pptx' library documentation for further graphic customization.
    prs = Presentation(template_path)
    # Add a new slide with a title and content
    slide_layout = prs.slide_layouts[0]  # Use the second layout in the template
    slide = prs.slides.add_slide(slide_layout)   
    # Add a title to the slide
    title_placeholder = slide.shapes.title
    title_placeholder.text = ppt_title    
         
    for i, outline in enumerate(outlines):
        lines=outline.split('\n')
        topic=lines[0]
        text='\n'.join(lines[1:])
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = topic
        title_shape = slide.shapes.title
        title_shape.text = topic

        # Format title font.
        title_shape.text_frame.paragraphs[0].font.bold = True
        title_shape.text_frame.paragraphs[0].font.size = Pt(42)
        left = Inches(1)
        top = Inches(2.5)
        width = Inches(9)
        height = Inches(4)

        # Add content to the slide
        content_placeholder = slide.placeholders[1]  # The first placeholder is the title
        content_placeholder.text = text
        

        if add_notes:
            notes_slide = slide.notes_slide
            text_frame = notes_slide.notes_text_frame
            text_frame.text = slide_notes[i]
            
    # Save presentation.
    pptx_filename = os.path.join(output_directory, f"{ppt_title}-{launguage}.pptx")
    prs.save(pptx_filename)     

# Define a function to ensure the conversation history does not exceed the maximum token limit
def ensure_under_token_limit(conversation, max_tokens):
    total_tokens = sum([len(message['content'].split()) for message in conversation])
    
    while total_tokens > max_tokens:
        # Remove the earliest messages until the total token count of the conversation history does not exceed the maximum token limit
        removed_message = conversation.pop(0)
        total_tokens -= len(removed_message['content'].split())
    
    return conversation

# Step 1: Generate the outline for the PowerPoint
conversation = [
    {"role": "system", "content": "You act as a cybersecurity expert."},
    {"role": "user", "content": f"Please draft a presentation outlines about {ppt_title} in {launguage} for a 20-minutes talk; in the presentation, please mention following key points: {additional_info}."}
]

outline_response = openai.ChatCompletion.create(
  model="gpt-3.5-turbo",
  messages=conversation
)

# Add the model's response to the conversation history
conversation.append({"role": "assistant", "content": outline_response['choices'][0]['message']['content']})

# Extract each page's topic from the outline
# Here it is assumed that each topic is on a separate line
outline = outline_response['choices'][0]['message']['content']
topics = outline.split('\n')
topics = topics[2:-1]

# First, combine all topics into a string with '\n' as the separator
topics_str = '\n'.join(topics)

# Then, split the string into multiple parts with empty lines (i.e., two consecutive '\n') as the separator
page_content = topics_str.split('\n\n')

# Finally, remove the leading and trailing white space (including '\n') from each part

outlines = [content.strip() for content in page_content]

print(f"presentation outline: {outlines}")

if step_by_step:
    # Generate PowerPoint with the outlines
    generate_pptx(outlines, ppt_title, template_path, output_directory)
    choice = input("Press 'C' to call OpenAI to generate presentation notes, or any other key to exit: ")
    if choice != 'C' and choice != 'c':
        sys.exit()


# Open the file to write the content
with open(f'ppt_files/{ppt_title}_ppt_notes-{launguage}.txt', 'w', encoding='utf8') as f:   
# Step 2: Generate content for each topic
    for i, topic in enumerate(outlines):
        conversation.append({"role": "user", "content": f"draft the PowerPoint notes in {launguage} in a spoken style to elaborate the slide that has the following content: {topic}"})

        # Ensure the conversation history does not exceed the maximum token limit
        conversation = ensure_under_token_limit(conversation, 3200)

        presentation_notes_response = openai.ChatCompletion.create(
        model="gpt-3.5-turbo",
        messages=conversation
        )
        time.sleep(1)

        # Add the model's response to the conversation history
        conversation.append({"role": "assistant", "content": presentation_notes_response['choices'][0]['message']['content']})

        presentation_notes = presentation_notes_response['choices'][0]['message']['content']
        slide_notes.append(presentation_notes)
        # Write the content to the file
        f.write(f"Slide {i+1} - {topic}: {presentation_notes}\n")        
        print(f"Slide {i+1} - {topic}: {presentation_notes}")
    

if step_by_step:
    choice = input("Press 'Y' to regenerate presentation ppt file with notes, or any other key to exit: ")
    if choice != 'Y' and choice != 'y':
        sys.exit()
    # Generate PowerPoint using outlines
generate_pptx(outlines, ppt_title, template_path, output_directory, add_notes=True)


