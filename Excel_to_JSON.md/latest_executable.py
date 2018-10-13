def excel_parsing(file, s_index=0):
    wb = xlrd.open_workbook(file)
    sheet = wb.sheet_by_index(s_index)
    data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
    questions = {}

    for row in range(sheet.nrows):
        if data[row][0]:
            if row not in questions:
                questions[row] = data[row][0]

    for row, question in questions.items():

        result_dict = {"QUESTION": question, "SUB-QUESTIONS": []}
        i = 0

        for sub_question_row in range(row + 1, sheet.nrows):
            if sub_question_row == sheet.nrows:
                break

            if data[sub_question_row][1]:
                sub_question = {"EACH-SUB-QUESTION": data[sub_question_row][1], "JURISDICTIONS": []}

                result_dict["SUB-QUESTIONS"].append(sub_question)
            else:
                break

            for jurisdiction_col_no in range(2, sheet.ncols):
                if data[sub_question_row][jurisdiction_col_no]:
                    jurisdiction_values = {"JURISDICTION": data[0][jurisdiction_col_no],
                                           "ANSWER": data[sub_question_row][jurisdiction_col_no]}
                else:
                    jurisdiction_values = {"JURISDICTION": data[0][jurisdiction_col_no],
                                           "ANSWER": ""}

                result_dict["SUB-QUESTIONS"][i]["JURISDICTIONS"].append(jurisdiction_values)

            i += 1

        json_mapping = json.dumps(result_dict)
        loaded_json = json.loads(json_mapping)
        print(loaded_json)


# file = '/Users/tanmaya/Documents/Hackathon-ChatBot/Leave Law State Law Comparison Report.xlsx'
file = '/Users/tanmaya/Documents/Hackathon-ChatBot/Q_andA_output_from_the_state_law_comparator_tool/Data Breach Notification State Law Comparison Report.xlsx'
excel_parsing(file, 0)
