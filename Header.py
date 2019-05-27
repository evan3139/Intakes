from docx import Document
from IsNumber import is_number


# This will create the excel sheet with all the formatting on it
def create_header(worksheet, worksheet_scores, file):
    contents = []
    header = []
    content = []
    score_header = []

    doc = Document(file)

    for line in doc.paragraphs:
        line.text = line.text.strip()
        if line.text == "":
            continue
        else:
            contents.append(line.text)

    for index, x in enumerate(contents):
        if ":" in x:
            headers, values = x.split(":")
            header.append(headers)
            content.append(values)
            if headers.lower() == "name":
                score_header.append(headers)
            elif "date" in x.lower():
                score_header.append(headers)
            elif "age" in x.lower():
                score_header.append(headers)
            elif "gender" in x.lower():
                score_header.append(headers)
            elif "race" in x.lower():
                score_header.append(headers)
            elif "bdi" in x.lower():
                score_header.append(headers)
            elif "ace" in x.lower():
                score_header.append(headers)
            elif "cage" in x.lower():
                score_header.append(headers)
            elif "bai" in x.lower():
                score_header.append(headers)
        else:
            header.append(x)

    # This will create the heading of the File for intake
    for index, x in enumerate(header):
        if is_number(header[index]):
            worksheet.write(0, index, int(header[index]))
        else:
            worksheet.write(0, index, header[index])

    for index, x in enumerate(score_header):
        if is_number(score_header[index]):
            worksheet_scores.write(0, index, int(x))
        else:
            worksheet_scores.write(0, index, x)
