from flask import Flask, request, redirect, url_for, flash, render_template, send_from_directory
from werkzeug.utils import secure_filename
from PIL import Image, ImageDraw, ImageFont
import pandas as pd
import os
import shutil
import img2pdf

app = Flask(__name__)

UPLOAD_FOLDER = r'uploads'
ALLOWED_EXTENSIONS = {'xml','xlsx'}

sheet_path=''

form_data={}

def allowed_file(filename):
	return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def middle(x1,x2,text,font):
	X_mid=(x1+x2)//2
	return X_mid- (font.getlength(text)//2)

def resizer(sign, SIGN_SIZE):
	x,y=sign.size
	X=x/SIGN_SIZE['x']
	Y=y/SIGN_SIZE['y']
	AR=x/y
	if X>Y:
		new_x=SIGN_SIZE['x']
		new_y=int(new_x/AR)
	else:
		new_y=SIGN_SIZE['y']
		new_x=int(AR * new_y )
	return sign.resize((new_x,new_y))

def center(x1,x2,x):
		length=x2-x1
		extra=length-x
		return x1+extra//2

def generator(FROM_DATE,TO_DATE,SIGNATURE_NAMES,DESIGNATIONS,text_2,text_3,template_given,text_2_size=36,text_3_size=36):
	#removing previous output
	for filename in os.listdir('output'):
		dir=os.path.join('output', filename)
		if os.path.isfile(dir):
			os.remove(dir)
	
	SHEET_PATH=r"uploads/sheet.xlsx"
	OUTPUT_PATH="output/"

	SIGN_PATHS=[
		"uploads/img1.png",
		"uploads/img2.png",
		"uploads/img3.png"
	]

	SIGN_BOX={"first":{'x1':90,'x2':405,'y1':660, 'y2':760},
			"second":{'x1':552,'x2': 868,'y1':660, 'y2':760 },
			"third":{'x1':1015,'x2':1333,'y1':660, 'y2':760 }}

	SIGN_SIZE= {'x':300, 'y':100}

	template_adjustment=0
	if TO_DATE:
		img= Image.open("template_with_qr.png")
	else:
		img=Image.open("template2_with_qr.png")
		template_adjustment=30
	
	if template_given==True:
		img=Image.open("uploads\\custom_template.png")

	sign1=Image.open(SIGN_PATHS[0])
	sign2=Image.open(SIGN_PATHS[1])
	sign3=Image.open(SIGN_PATHS[2])


	#pasting signs
	sign1=resizer(sign1,SIGN_SIZE)
	extra=SIGN_SIZE['y']-sign1.size[1]
	img.paste(sign1,(center(SIGN_BOX['first']['x1'],SIGN_BOX['first']['x2'],sign1.size[0]),SIGN_BOX['first']['y1']+extra),mask=sign1)

	sign2=resizer(sign2,SIGN_SIZE)
	extra=SIGN_SIZE['y']-sign2.size[1]
	img.paste(sign2,(center(SIGN_BOX['second']['x1'],SIGN_BOX['second']['x2'],sign2.size[0]),SIGN_BOX['second']['y1']+extra),mask=sign2)

	sign3=resizer(sign3,SIGN_SIZE)
	extra=SIGN_SIZE['y']-sign3.size[1]
	img.paste(sign3,(center(SIGN_BOX['third']['x1'],SIGN_BOX['third']['x2'],sign3.size[0]),SIGN_BOX['third']['y1']+extra),mask=sign3)

	draw_obj= ImageDraw.Draw(img)

	arial=ImageFont.truetype("fonts\ARIALBD.TTF", size=36)
	#small_arial=ImageFont.truetype(r"E:\certificate generator 2.0\fonts\ARIALBD.TTF", size=30)
	segoe_UI_bold=ImageFont.truetype(r"fonts\Segoe UI Bold.ttf",size=30)
	arial_normal=ImageFont.truetype(r"fonts\ARIAL.TTF",size=28)

	# text_2="Summer School Program 2024"
	# text_3="Center for Artificial Intelligence"

	#2nd line
	if text_2_size!=36:
		arial_2=ImageFont.truetype(r"fonts\ARIALBD.TTF", size=text_2_size)
	else:
		arial_2=arial
	y_extra_line2=arial.getbbox(text_2)[3]-arial_2.getbbox(text_2)[3]
	draw_obj.text((middle(340,1330,text_2,arial_2),465+y_extra_line2), text_2, font=arial_2, fill=(0,0,0))

	#3rd line
	if text_3_size!=36:
		arial_3=ImageFont.truetype(r"fonts\ARIALBD.TTF", size=text_3_size)
	else:
		arial_3=arial
	y_extra_line3=arial.getbbox(text_3)[3]-arial_3.getbbox(text_3)[3]
	draw_obj.text((middle(270,1330,text_3,arial_3),525+y_extra_line3), text_3, font=arial_3, fill=(0,0,0))

	#date line
	draw_obj.text((155-template_adjustment,595), FROM_DATE , font=arial, fill=(0,0,0))
	draw_obj.text((355,595), TO_DATE , font=arial, fill=(0,0,0))

	# signature names
	for i,j in enumerate(SIGN_BOX.keys()):
		draw_obj.text((middle(SIGN_BOX[j]['x1'],SIGN_BOX[j]['x2'],SIGNATURE_NAMES[i],segoe_UI_bold),760)
					,SIGNATURE_NAMES[i]
					,font=segoe_UI_bold, 
					fill=(0,0,0))
		
	#designation 
	for i,j in enumerate(SIGN_BOX.keys()):
		draw_obj.text((middle(SIGN_BOX[j]['x1'],SIGN_BOX[j]['x2'],DESIGNATIONS[i],arial_normal),800)
					,DESIGNATIONS[i]
					,font=arial_normal, 
					fill=(0,0,0))
		
	sheet_data = pd.read_excel(SHEET_PATH)

	for i,j in sheet_data.loc[:,["Name","Optional"]].values:
		half_output=img.copy()
		obj=ImageDraw.Draw(half_output)
		if str(j)=='nan':
			text=i
		else:
			text=i+f"({j})"
		obj.text((middle(560,1330,text,arial),410), text, font=arial, fill=(0,0,0))
		half_output.save("temp.png")
		pdf_bytes=img2pdf.convert("temp.png")
		file=open(os.path.join(OUTPUT_PATH+f"{i}.pdf"),"wb")
		file.write(pdf_bytes)
		file.close()
		
        # BELOW METHOD THERE WAS A LOT OF QUALITY LOSS WHILE CONVERTING TO PDF
		# temp=half_output.convert('RGB')
		# temp.save(os.path.join(OUTPUT_PATH+f"{i}.pdf"))

	shutil.make_archive("outputs/output","zip","output")



@app.route('/', methods=['GET', 'POST'])
def upload():
	global sheet_path, form_data
	if request.method == 'POST':
		template_given=False
		from_date=request.form.get('from')
		to_date=request.form.get('to')

		#changing format of date
		from_date='.'.join(from_date.split('-')[::-1])
		to_date='.'.join(to_date.split('-')[::-1])
		from_date=from_date[:-4]+from_date[-2:]
		to_date=to_date[:-4]+to_date[-2:]


		form_data={'from':from_date, 'to':to_date}
		print(form_data)
		file = request.files['sheet']
		img1=request.files['sign1_img']
		img2=request.files['sign2_img']
		img3=request.files['sign3_img']

		img_lst=[img1,img2,img3]

		if not file or not img1 or not img2 or not img3:
			return render_template('index.html', msg="Something went wrong... Please try again!")
		
		# uploading and saving sheets
		sheet_path=os.path.join(UPLOAD_FOLDER, 'sheet.'+file.filename.rsplit('.')[-1])
		file.save(sheet_path)
		
		#uploading and saving signatures
		for i,j in enumerate(img_lst):
			img_path=os.path.join(UPLOAD_FOLDER, f"img{i+1}.png")
			j.save(img_path)
		
		#saving template if given
		custom_template=request.files['template']
		if custom_template:
			custom_template.save(os.path.join(UPLOAD_FOLDER,"custom_template.png"))
			template_given=True

		# getting names
		SIGN_NAMES=[request.form.get('sign1_name'),request.form.get('sign2_name'),request.form.get('sign3_name')]
		SIGN_DESIG=[request.form.get('sign1_desig'),request.form.get('sign2_desig'),request.form.get('sign3_desig')]
		
		#getting text
		text_2=request.form.get('2nd_line')
		text_3=request.form.get('3rd_line')
		
		font_size2=request.form.get('2nd_line_font_size')
		font_size3=request.form.get("3rd_line_font_size")

		generator(from_date,to_date,SIGN_NAMES,SIGN_DESIG,text_2,text_3,
			template_given,
			int(font_size2),int(font_size3))

		return redirect('/download') #redirecting will perform a GET request 
	return render_template('index.html')

@app.route('/download', methods=['POST','GET'])
def download():
	return render_template('download_page.html')

@app.route("/downloading")
def downloading():
	return send_from_directory("outputs","output.zip")


if __name__ == '__main__':
	app.run(debug=True)
