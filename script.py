import pandas as pd

# Load your Excel file into a DataFrame
df = pd.read_excel('planilha.xlsx')

from PIL import Image, ImageDraw, ImageFont

# Define A4 paper dimensions (in pixels)
a4_width, a4_height = 2550, 3300  # Standard A4 size at 300 DPI

top_margin = 150

# Set up card dimensions
card_width, card_height = a4_width // 2, (a4_height - 2*top_margin) // 5  # Adjust based on your label size

name_size = 80
inst_size = 60

margin = card_height - (name_size + inst_size)
card_y_offset = margin // 2

# Helper function to draw centered text
def draw_centered_text(draw, text, font, y):
    text_width = draw.textlength(text, font=font)
    x = (card_width - text_width) / 2  # Calculate x to center the text
    draw.text((x, y), text, fill='black', font=font)

def create_id_card(data):
    # Create a blank image for the card
    card = Image.new('RGB', (card_width, card_height), 'white')
    draw = ImageDraw.Draw(card)
    
    # Load font (make sure to adjust path or choose a different font)
    font_name = ImageFont.truetype('LiberationSerif-Regular.ttf', name_size)
    font_inst = ImageFont.truetype('LiberationSerif-Regular.ttf', inst_size)

    # Draw text on the card (adjust positions as needed)
    draw_centered_text(draw, f"{data['Nome']}", font=font_name, y=card_y_offset)
    draw_centered_text(draw, f"{data['Instituição']}", font=font_inst, y=name_size + card_y_offset)

    return card

for index, row in df.iterrows():
    id_card = create_id_card(row)
    id_card.save(f"id_card_{index}.png")  # Save each card as an image

a4_sheet = Image.new('RGB', (a4_width, a4_height), 'white')

page_number = 1
x_offset, y_offset = 0, top_margin
margin = 0  # Margin between cards

for index, row in df.iterrows():
    id_card = create_id_card(row)

    if x_offset + card_width > a4_width:
        x_offset = 0
        y_offset += card_height + margin

    if y_offset + card_height > a4_height:
        # Save the current page and start a new one
        a4_sheet.save(f'a4_sheet_page_{page_number}.png')
        page_number += 1
        a4_sheet = Image.new('RGB', (a4_width, a4_height), 'white')  # New A4 sheet
        x_offset, y_offset = 0, top_margin  # Reset offsets for new page

    a4_sheet.paste(id_card, (x_offset, y_offset))
    x_offset += card_width + margin

# Save the last page if not empty
if x_offset != 0 or y_offset != 0:
    a4_sheet.save(f'a4_sheet_page_{page_number}.png')
