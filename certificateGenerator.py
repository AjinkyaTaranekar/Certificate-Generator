from PIL import Image
from PIL import ImageFont
from PIL import ImageDraw
from google.cloud import vision
from google.cloud.vision import types
import io

def assemble_word(word):
    assembled_word = ""
    for symbol in word.symbols:
        assembled_word += symbol.text
    return assembled_word

def find_word_location(document, word_to_find):
	
    for page in document.pages:
        for block in page.blocks:
            for paragraph in block.paragraphs:
                for word in paragraph.words:
                    assembled_word = assemble_word(word)
                    if (assembled_word == word_to_find):
 		   	return word.bounding_box


vision_client = vision.ImageAnnotatorClient()

with io.open('certificate.png', 'rb') as image_file:
    content = image_file.read()

content_image = types.Image(content=content)
response = vision_client.document_text_detection(image=content_image)
document = response.full_text_annotation

location=(find_word_location(document, "that"))
img = Image.open('certificate.png')
draw = ImageDraw.Draw(img)
font = ImageFont.truetype("/usr/share/fonts/truetype/freefont/FreeMono.ttf", 8, encoding="unic")

draw.text((location.vertices[1].x +5,location.vertices[1].y-1), 'Ajinkya Taranekar', (0,0,0), font=font)
img.save('out_file.png')

location=(find_word_location(document, "on"))

draw = ImageDraw.Draw(img)
font = ImageFont.truetype("/usr/share/fonts/truetype/freefont/FreeMono.ttf", 12, encoding="unic")
draw.text((location.vertices[1].x +5,location.vertices[1].y -1), '12/10/2019', (0,0,0), font=font)
img.save('out_file.png')

location =find_word_location(document, "at")
draw = ImageDraw.Draw(img)
font = ImageFont.truetype("/usr/share/fonts/truetype/freefont/FreeMono.ttf", 12, encoding="unic")
draw.text((location.vertices[1].x +5,location.vertices[1].y+20), 'Bhopal', (0,0,0), font=font)
img.save('out_file.png')

location=find_word_location(document, "entitled")

font = ImageFont.truetype("/usr/share/fonts/truetype/freefont/FreeMono.ttf", 12, encoding="unic")
draw = ImageDraw.Draw(img)
draw.text((location.vertices[1].x +5,location.vertices[1].y -1), "Version Beta", (0,0,0), font=font)
img.save('out_file.png')
