from PIL import Image
from PIL import ImageFont
from PIL import ImageDraw
from google.cloud import vision
from google.cloud.vision import types
import io
import xlrd 
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders,email
import email.mime.application

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
def findEmailCol():
	for i in range(sheet.ncols):
		if(str(sheet.cell_value(0,i)).lower()=="email"):
			return i


vision_client = vision.ImageAnnotatorClient()

image_file='certificate.png'
with io.open(image_file, 'rb') as image_file2:
    content = image_file2.read()

content_image = types.Image(content=content)
response = vision_client.document_text_detection(image=content_image)
document = response.full_text_annotation

font = ImageFont.truetype("/usr/share/fonts/truetype/freefont/FreeMono.ttf", 8, encoding="unic")
mail_content = '''Hello,
Hurray You have won the Version Beta.
Let's Celebrate.
Party by Abhi Jain.
Thank You
'''
#The mail addresses and password
sender_address = 'gsbeta2k19@gmail.com'
sender_pass = 'versionbeta2k19'

#Setup the MIME
message = MIMEMultipart()
message['From'] = sender_address
message['Subject'] = 'Winner of Version Beta'
#The subject line
#The body and the attachments for the mail

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

	attach_file=0
	receiver_address = str(sheet.cell_value(i,findEmailCol()))

	message['To'] = receiver_address
	attach_file_name = 'Congrats '+str(sheet.cell_value(i,findNameCol()))+'.png' 
	attach_file = open(output_image,'rb')

	payload = MIMEBase('application', 'octate-stream')
	payload.set_payload((attach_file).read())

#add payload header with filename
	payload.add_header('Content-Decomposition', 'attachment', filename=attach_file_name)
	message.attach(payload)
	message.attach(MIMEText(mail_content, 'plain'))

#Create SMTP session for sending the mail
	session = smtplib.SMTP('smtp.gmail.com', 587) #use gmail with port
	session.starttls() #enable security
	session.login(sender_address, sender_pass) #login with mail_id and password
	text = message.as_string()
	session.sendmail(sender_address, receiver_address, text)
	session.quit()

