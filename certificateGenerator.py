from PIL import Image
from PIL import ImageFont
from PIL import ImageDraw
from google.cloud import vision
from google.cloud.vision import types
import io
import xlrd 
  
# Give the location of the file 
loc = ("sample.xlsx") 
  
wb = xlrd.open_workbook(loc) 
sheet = wb.sheet_by_index(0) 
sheet.cell_value(0, 0) 
  
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

def findNameCol():
	for i in range(sheet.ncols):
		if(str(sheet.cell_value(0,i)).lower()=="name"):
			return i


vision_client = vision.ImageAnnotatorClient()

image_file='certificate.png'
with io.open(image_file, 'rb') as image_file2:
    content = image_file2.read()

content_image = types.Image(content=content)
response = vision_client.document_text_detection(image=content_image)
document = response.full_text_annotation

font = ImageFont.truetype("/usr/share/fonts/truetype/freefont/FreeMono.ttf", 8, encoding="unic")

for i in range(1,sheet.nrows): 
	location=(find_word_location(document, "that"))
	img = Image.open(image_file)
	draw = ImageDraw.Draw(img)	
	draw.text((location.vertices[1].x +5,location.vertices[1].y-1), str(sheet.cell_value(i,findNameCol())), (0,0,0), font=font)
	
	location=(find_word_location(document, "on"))
	draw = ImageDraw.Draw(img)
	draw.text((location.vertices[1].x +5,location.vertices[1].y -1), '12/10/2019', (0,0,0), font=font)

	location =find_word_location(document, "at")
	draw = ImageDraw.Draw(img)
	draw.text((location.vertices[1].x +5,location.vertices[1].y+20), 'Bhopal', (0,0,0), font=font)

	location=find_word_location(document, "entitled")
	draw = ImageDraw.Draw(img)
	draw.text((location.vertices[1].x +5,location.vertices[1].y -1), "Version Beta", (0,0,0), font=font)
	output_image='out_file'+str(i)+'.png'	
	img.save(output_image)
