from fpdf import FPDF,XPos,YPos

pdf=FPDF(orientation='P',format='A4',unit='pt')
#this creates a fromat kind of describing the pdf 

pdf.add_page()
#need to add a page in order to see the pdf

pdf.image('tiger.jpeg',w=80,h=60)
#add an image using pdf.image() that accepts the path of image, wigth and height of that image

pdf.set_font(family='times',style='B',size=20)
pdf.cell(w=0, h=50,text="Bengal Tiger", align='C',new_x=XPos.LMARGIN,new_y=YPos.NEXT)
'''for every cell declaration to add text in the pdf, we need to first describe the pdf.set_font(), then 
    procced with pdf.cell() add the width=0 to have the total row assigned for the cell
    - to create the line breaks add ln=1
    - to add the text towards the left margin and place the next cell text to new line use 
        new_x=XPos.LMARGIN and new_y=YPos.NEXT'''

pdf.set_font(family='times',style='B',size=20)
pdf.cell(w=0,h=15,text="Description",new_x=XPos.LMARGIN,new_y=YPos.NEXT)

pdf.set_font(family='times',size=12)
txt1='''The Bengal tiger is a population of the Panthera tigris tigris subspecies. 
It ranks among the largest of wild cats. It is distributed from India, southern Nepal, Bangladesh, Bhutan to Southwestern China. 
Its historical range extended to the Indus Basin until the early 19th century, and it is thought to have been present in the Indian subcontinent since the 
Late Pleistocene about 12,000 to 16,500 years ago. It is threatened by poaching, habitat loss and habitat fragmentation. It is also National animal of India'''
pdf.multi_cell(w=0,h=25,text=txt1,new_x=XPos.LMARGIN,new_y=YPos.NEXT)
#to have the cells close enoudh we need to det the [h=""] 
# and add the multiline cell using multi_cell() function

pdf.set_font(family='times',style='B',size=12)
pdf.cell(w=100,h=25,text='Kingdom:')

pdf.set_font(family='times',size=12)
pdf.cell(w=100,h=25,text='Animalia',new_x=XPos.LMARGIN,new_y=YPos.NEXT)

pdf.set_font(family='times',style='B',size=12)
pdf.cell(w=100,h=25,text='Phylum:')

pdf.set_font(family='times',size=12)
pdf.cell(w=100,h=25,text='Chordata')

pdf.output('Tiger.pdf')

