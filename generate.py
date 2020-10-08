from PIL import Image,ImageDraw,ImageFont
import matplotlib.pyplot as plt
import  pandas as pd
from termcolor import colored
from tqdm import tqdm_gui
import os 
from datetime import datetime
import shutil


# Usually a .csv file but here i've used .xlsx file
# Make sure to have 'Name' column in it
file = pd.read_excel('event.xlsx')

# Uni name is hard coded for CUI only, you can uncomment below option for making it generic
# uniName = input("Enter University Name: ")
uniName = "VIT-AP University"

# uniAcronym = input("Enter University Acronym: ")
uniAcronym = "CSI VIT-AP"


eventName = input("Enter Event Name: ")
# leadName = input("Lead Name: ")
leadName = "Mervin Joseph"

currentDate = datetime.date(datetime.now())

# where certificates will be generated
fname = 'certificates/'
if os.path.exists(fname):
    shutil.rmtree(fname)
os.mkdir(fname)

# Getting 'Name' columns from 'file'
for names in tqdm_gui(file['Name']):
    image = Image.new('RGB',(1920,1080),(255,255,255))

    draw = ImageDraw.Draw(image)
    font_path = './Almondita.ttf'

    # font of csi
    fontdev = ImageFont.truetype('arial.ttf', size=35)
    # font of title
    fonttitle = ImageFont.truetype('arialbd.ttf', size=55)
    # font of Certificate
    fontcert = ImageFont.truetype('Pacifico.ttf', size=55)
    # font of participant name
    fontname = ImageFont.truetype('arial.ttf', size=35)
    # font of signature
    signature = ImageFont.truetype(font_path, 150)

    # colors for various writings
    colordev = 'rgb(128, 128, 128)'
    colorcert = 'rgb(89, 89, 89)'
    colorname = 'rgb(77, 148, 255)'
    colortitle = 'rgb(86, 101, 115)'
    colorNamecsiLead = 'rgb(229, 57, 53)'


    # csi logo
    csi_logo = Image.open('logo.jpg')
    # Resizing the csi Logo
    csi_logo = csi_logo.resize((175,175))

    # Left Side style Image
    side_style = Image.open('leftSide.png')

    # Putting Logo & Left style Image
    image.paste(csi_logo, (1700,10))
    image.paste(side_style, (0,0))


    participation_message = f"is hereby awarded this Certificate of Participation on successfully attending \n{eventName} at {uniName} organized by {uniAcronym}."

    draw.text((1050,85),'CSI CHAPTER VIT-AP',font=fonttitle,fill=colorcert)
    draw.text((950,200),'Certificate of Participation',font=fontcert,fill=colorcert)
    draw.text((950,325), eventName + ' Participant',font=ImageFont.truetype('arial.ttf', size=32), fill=colordev)
    draw.text((950,400),names,font=fonttitle,fill=colorname)
    draw.text((950,500),participation_message,font=ImageFont.truetype('arial.ttf', size=25),fill='rgb(102, 102, 102)')
    draw.text((960,650),leadName,font=signature,fill=colorNamecsiLead)
    draw.line((950,800, 1520,800), fill='rgb(128, 128, 128)', width=3)

    draw.text((950,820),f'Computer Society of India, {uniAcronym}',font=ImageFont.truetype('./arialbd.ttf', size=22),fill=colorcert)
    draw.text((950,920), 'Signing Authority', font=ImageFont.truetype('arial.ttf', size=22), fill='rgb(128,128,128)')
    draw.text((950,960),  'President ' + uniAcronym , font=ImageFont.truetype('./arialbd.ttf', size=25), fill='rgb(128,128,128)')
    draw.text((1640,950),'Certificate ID:',font=ImageFont.truetype('arial.ttf', size=20),fill=colorcert)
    draw.text((1640,980),f'Issue Date: {currentDate}',font=ImageFont.truetype('arial.ttf', size=20),fill=colorcert)
    image.save('certificates/'+names+'.png')

