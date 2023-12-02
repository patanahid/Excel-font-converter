from flask import Flask, render_template, request, send_file
from openpyxl import load_workbook
from openpyxl.styles import Font
import os

app = Flask(__name__)

def convert_kruti_to_unicode(modified_substring):
    global array_one, array_two, array_one_length,text_size,processed_text,sthiti1,sthiti2,chale_chalo,max_text_size
    array_one = ["ñ","Q+Z","sas","aa",")Z","ZZ","‘","’","“","”","å","ƒ","„","…","†","‡","ˆ","‰","Š","‹","¶+","d+","[+k","[+","x+","T+","t+","M+","<+","Q+",";+","j+","u+","Ùk","Ù","ä","–","—","é","™","=kk","f=k","à","á","â","ã","ºz","º","í","{k","{","=","«","Nî","Vî","Bî","Mî","<î","|","K","}","J","Vª","Mª","<ªª","Nª","Ø","Ý","nzZ","æ","ç","Á","xz","#",":","v‚","vks","vkS","vk","v","b±","Ã","bZ","b","m","Å",",s",",","_","ô","d","Dk","D","[k","[","x","Xk","X","Ä","?k","?","³","pkS","p","Pk","P","N","t","Tk","T",">","÷","¥","ê","ë","V","B","ì","ï","M+","<+","M","<",".k",".","r","Rk","R","Fk","F",")","n","/k","èk","/","Ë","è","u","Uk","U","i","Ik","I","Q","¶","c","Ck","C","Hk","H","e","Ek","E",";","¸","j","y","Yk","Y","G","o","Ok","O","'k","'","\"k","\"","l","Lk","L","g","È","z","Ì","Í","Î","Ï","Ñ","Ò","Ó","Ô","Ö","Ø","Ù","Ük","Ü","‚","ks","kS","k","h","q","w","`","s","S","a","¡","%","W","•","·","∙","·","~j","~","\\","+"," ः","^","*","Þ","ß","(","¼","½","¿","À","¾","A","-","&","&","Œ","]","~ ","@"]
    array_two = ["॰","QZ+","sa","a","र्द्ध","Z","\"","\"","'","'","०","१","२","३","४","५","६","७","८","९","फ़्","क़","ख़","ख़्","ग़","ज़्","ज़","ड़","ढ़","फ़","य़","ऱ","ऩ","त्त","त्त्","क्त","दृ","कृ","न्न","न्न्","=k","f=","ह्न","ह्य","हृ","ह्म","ह्र","ह्","द्द","क्ष","क्ष्","त्र","त्र्","छ्य","ट्य","ठ्य","ड्य","ढ्य","द्य","ज्ञ","द्व","श्र","ट्र","ड्र","ढ्र","छ्र","क्र","फ्र","र्द्र","द्र","प्र","प्र","ग्र","रु","रू","ऑ","ओ","औ","आ","अ","ईं","ई","ई","इ","उ","ऊ","ऐ","ए","ऋ","क्क","क","क","क्","ख","ख्","ग","ग","ग्","घ","घ","घ्","ङ","चै","च","च","च्","छ","ज","ज","ज्","झ","झ्","ञ","ट्ट","ट्ठ","ट","ठ","ड्ड","ड्ढ","ड़","ढ़","ड","ढ","ण","ण्","त","त","त्","थ","थ्","द्ध","द","ध","ध","ध्","ध्","ध्","न","न","न्","प","प","प्","फ","फ्","ब","ब","ब्","भ","भ्","म","म","म्","य","य्","र","ल","ल","ल्","ळ","व","व","व्","श","श्","ष","ष्","स","स","स्","ह","ीं","्र","द्द","ट्ट","ट्ठ","ड्ड","कृ","भ","्य","ड्ढ","झ्","क्र","त्त्","श","श्","ॉ","ो","ौ","ा","ी","ु","ू","ृ","े","ै","ं","ँ","ः","ॅ","ऽ","ऽ","ऽ","ऽ","्र","्","?","़",":","‘","’","“","”",";","(",")","{","}","=","।",".","-","µ","॰",",","् "]
    array_one_length = len(array_one)
    
    text_size = len(str(modified_substring))
    processed_text = ''
    sthiti1 = 0
    sthiti2 = 0
    chale_chalo = 1
    max_text_size = 6000
    
    while chale_chalo == 1:
        sthiti1 = sthiti2
        if sthiti2 < (text_size - max_text_size):
            sthiti2 += max_text_size
            while modified_substring[sthiti2] != ' ':
                sthiti2 -= 1
        else:
            sthiti2 = text_size
            chale_chalo = 0
        
        modified_substring = modified_substring[sthiti1:sthiti2]


        if modified_substring:
            for input_symbol_idx in range(len(array_one)):
                idx = 0
                try:
                    while idx != -1:
                        modified_substring = modified_substring.replace(array_one[input_symbol_idx], array_two[input_symbol_idx])
                        
                        idx = modified_substring.find(array_one[input_symbol_idx])
                except:
                    pass
        
            modified_substring = modified_substring.replace("±", "Zं")
            modified_substring = modified_substring.replace("Æ", "र्f")
            
            position_of_i = modified_substring.find("f")
            while position_of_i != -1:
                charecter_next_to_i = modified_substring[position_of_i + 1]
                charecter_to_be_replaced = "f" + charecter_next_to_i
                modified_substring = modified_substring.replace(charecter_to_be_replaced, charecter_next_to_i + "ि")
                position_of_i = modified_substring.find("f", position_of_i + 1)
            
            # modified_substring = modified_substring.replace("Ç", "fa")
            
            position_of_i = modified_substring.find("Ç")
            while position_of_i != -1:
                charecter_next_to_ip2 = modified_substring[position_of_i + 2]
                charecter_to_be_replaced = "Ç" + charecter_next_to_ip2
                modified_substring = modified_substring.replace(charecter_to_be_replaced, charecter_next_to_ip2 + "िं")
                position_of_i = modified_substring.find("Ç", position_of_i + 2)
            
            modified_substring = modified_substring.replace("Ê", "ीZ")
            
            position_of_wrong_ee = modified_substring.find("ि्")
            while position_of_wrong_ee != -1:
                consonent_next_to_wrong_ee = modified_substring[position_of_wrong_ee + 2]
                charecter_to_be_replaced = "ि्" + consonent_next_to_wrong_ee
                modified_substring = modified_substring.replace(charecter_to_be_replaced, "्" + consonent_next_to_wrong_ee + "ि")
                position_of_wrong_ee = modified_substring.find("ि्", position_of_wrong_ee + 2)
            
            set_of_matras = "अ आ इ ई उ ऊ ए ऐ ओ औ ा ि ी ु ू ृ े ै ो ौ ं : ँ ॅ"
            position_of_R = modified_substring.find("Z")
            while position_of_R > 0:
                probable_position_of_half_r = position_of_R - 1
                charecter_at_probable_position_of_half_r = modified_substring[probable_position_of_half_r]
                while set_of_matras.find(charecter_at_probable_position_of_half_r) != -1:
                    probable_position_of_half_r -= 1
                    charecter_at_probable_position_of_half_r = modified_substring[probable_position_of_half_r]
                charecter_to_be_replaced = modified_substring[probable_position_of_half_r:(position_of_R - probable_position_of_half_r)]
                new_replacement_string = "र्" + charecter_to_be_replaced
                charecter_to_be_replaced = charecter_to_be_replaced + "Z"
                modified_substring = modified_substring.replace('Z', 'र्')
                position_of_R = modified_substring.find("Z")


        
        processed_text += modified_substring
        return processed_text
    
    
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process_file():
    # Assuming your HTML file input has the name 'file'
    uploaded_file = request.files['file']
    
    # Save the uploaded file
    uploaded_file.save("static/input.xlsx")

    # Indicate processing to the user
    os.rename("static/input.xlsx", "static/input_in_progress.xlsx")

    wb = load_workbook(filename="static/input_in_progress.xlsx")

    # Rest of your code for processing the Excel file
    sheets_to_duplicate = []

    Original_sheets = wb.sheetnames 

    for sheet_name in wb.sheetnames:
        sheets_to_duplicate.append(sheet_name)

    for sheet_name in sheets_to_duplicate:
        source = wb[sheet_name]
        target = wb.copy_worksheet(source)

        for row in target.iter_rows():
            for cell in row:
                if "devlys".upper() in cell.font.name.upper() or "kruti".upper() in cell.font.name.upper():
                    if cell.value is not None:
                        cell.value = convert_kruti_to_unicode(str(cell.value))
                        cell.font = Font(name="Calibri", size=cell.font.size, bold=cell.font.bold,
                                        italic=cell.font.italic, strikethrough=cell.font.strikethrough,
                                        underline=cell.font.underline, strike=cell.font.strike,
                                        color=cell.font.color, vertAlign=cell.font.vertAlign)

    for sheet in Original_sheets:
        del wb[sheet]

    # Save the modified workbook as output
    output_filename = "static/" + os.path.splitext(uploaded_file.filename)[0] + "-converted.xlsx"
    wb.save(output_filename)

    # Remove the in-progress indicator
    os.rename("static/input_in_progress.xlsx", "static/input.xlsx")

    # Provide a download link for the processed file
    return send_file(output_filename, as_attachment=True, download_name="output.xlsx")

if __name__ == '__main__':
    app.run(debug=True)
