from docx import Document
from IsNumber import is_number


def fill_sheets(worksheet, sheet, filename, row):
    doc = Document(filename)
    inputs = []
    data = []
    score_data = []

    for line in doc.paragraphs:
        line.text = line.text.strip()
        if line.text == "":
            continue
        else:
            inputs.append(line.text)

    for index, i in enumerate(inputs):
        if ":" in i:
            useless, values = i.split(":")
            data.append(values)
            if useless.lower() == "name":
                score_data.append(values)
            elif "date" in i.lower():
                score_data.append(values)
            elif "age" in i.lower():
                score_data.append(values)
            elif "gender" in i.lower():
                score_data.append(values)
            elif "race" in i.lower():
                score_data.append(values)
            elif "bdi" in i.lower():
                score_data.append(values)
            elif "ace" in i.lower():
                score_data.append(values)
            elif "cage" in i.lower():
                score_data.append(values)
            elif "bai" in i.lower():
                score_data.append(values)
        else:
            data.append(i)

    data = [x.strip() for x in data]
    score_data = [x.strip() for x in score_data]
    data[1] = data[1].replace(" ", "")

    for index, x in enumerate(data):
        if is_number(x):
            worksheet.write(row, index, int(x))
        else:
            worksheet.write(row, index, x)

    for index, x in enumerate(score_data):
        if is_number(x):
            sheet.write(row, index, int(x))
        else:
            sheet.write(row, index, x)