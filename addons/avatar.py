from PIL import Image, ImageDraw, ImageFont

def create_number_image(number, filename="output.png", font_color="white", bg_color="black"):
    """
    Создает картинку 256x256 с числом строго по центру.
    """

    size = (256, 256)
    image = Image.new("RGB", size, color=bg_color)
    draw = ImageDraw.Draw(image)

    try:
        font = ImageFont.truetype("arial.ttf", size=150)
    except IOError:
        font = ImageFont.load_default()

    text = str(number)
    
    bbox = draw.textbbox((0, 0), text, font=font)
    text_width = bbox[2] - bbox[0]
    text_height = bbox[3] - bbox[1]
    
    x = (size[0] - text_width) // 2 - bbox[0]
    y = (size[1] - text_height) // 2 - bbox[1]

    draw.text((x, y), text, fill=font_color, font=font)
    
def save_image(image, filename):
    image.save(filename)
    
# Example:
# create_number_image(number=2, filename="number.png", font_color="black", bg_color="white")
