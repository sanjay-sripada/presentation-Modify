import aspose.slides as slides
import re
import aspose.pydrawing as drawing
import aspose.slides.charts as charts # type: ignore


presentation= slides.Presentation('assignment.pptx')
slide_1=presentation.slides[0]

pattern =re.compile(r'0[1-9]|1[0-9]|2[0-6]')
mapping = {i: chr(64 + i) for i in range(1, 27)}

#changes image to car 
with open('car.jpg','rb') as image_file:
    new_image= presentation.images.add_image(image_file)

for shape in slide_1.shapes:
    if isinstance(shape,slides.PictureFrame):
        shape.picture_format.picture.image=new_image
# change number to alphabets 
    if isinstance(shape,slides.AutoShape):
        for paragraphs in shape.text_frame.paragraphs:
            if pattern.match(paragraphs.text):
                number= int(paragraphs.text)
                if number in mapping:
                    paragraphs.text=mapping[number]
                                    
# change slide 2 green header to  blue color 
slide_2= presentation.slides[1]
for shape in slide_2.shapes:
    if isinstance(shape,slides.Table):
        table = shape
        header_row=table.rows[0]
        for col_idx in range(1,7):
            cell = header_row[col_idx]
            cell.cell_format.fill_format.fill_type = slides.FillType.SOLID
            cell.cell_format.fill_format.solid_fill_color.color = drawing.Color.blue
            
# change slide 3 title color from Black to blue 
slide_3= presentation.slides[2]
for shape in slide_3.shapes:
    if isinstance(shape,slides.AutoShape):
        for paragraphs in shape.text_frame.paragraphs:
            for portion in paragraphs.portions:
                if(portion.portion_format.font_height > 20):
                    portion.portion_format.fill_format.fill_type = slides.FillType.SOLID
                    portion.portion_format.fill_format.solid_fill_color.color = drawing.Color.blue
 # Change  char color from Blue to red                   
    if isinstance(shape,charts.Chart):
        for series in shape.chart_data.series:
            for point in series.data_points:
                point.format.fill.fill_type = slides.FillType.SOLID
                point.format.fill.solid_fill_color.color = drawing.Color.red
            
        
presentation.save('modif.pptx',slides.export.SaveFormat.PPTX)
    