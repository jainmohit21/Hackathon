import xlrd
import json

def excel_parsing(file, s_index=0):

    wb = xlrd.open_workbook(file)
    sheet = wb.sheet_by_index(s_index)
    # getting the data from the worksheet
    data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
    for row in range(sheet.nrows):
        if data[row][0]:
            result_dict = {"QUESTION": data[row][0], "SUB-QUESTIONS": []}
            i = 0
            for sub_question_row in range(row+1, row+7):
                if sub_question_row == sheet.nrows:
                    break

                if data[sub_question_row][1]:
                    sub_question = {"EACH-SUB-QUESTION": data[sub_question_row][1], "JURISDICTIONS": []}

                result_dict["SUB-QUESTIONS"].append(sub_question)

                for jurisdiction_col_no in range(2, sheet.ncols):

                    jurisdiction_values = {"JURISDICTION": data[0][jurisdiction_col_no], "ANSWER": data[sub_question_row][jurisdiction_col_no]}

                    result_dict["SUB-QUESTIONS"][i]["JURISDICTIONS"].append(jurisdiction_values)
                i += 1
            j = i
            json_mapping = json.dumps(result_dict)
            loaded_json = json.loads(json_mapping)
            print(loaded_json)


file = 'C:\Users\Anshul Chaudhary\Desktop\Leave Law State Law Comparison Report.xlsx'
excel_parsing(file, 0)
