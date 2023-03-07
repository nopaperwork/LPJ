import openpyxl
import sqlite3 as sl
con = sl.connect('dont_touch.db')

tracker_dict ={}
# from excel to DB
book = openpyxl.load_workbook('output_db.xlsx')
sheet = book['Sheet1']
# here you iterate over the rows in the specific column
for row in range(2, sheet.max_row + 1):
    main_key = red_id  = destination_pin_code = addressee = addressee_address = None
    for column in "AGHIJL":  # Here you can add or reduce the columns
        cell_name = "{}{}".format(column, row)
        if column == "A":
            main_key = sheet[cell_name].value
        if column == "G":
            weight = sheet[cell_name].value
            if weight == None :
                weight = ''
        if column == "H":
            net_value = sheet[cell_name].value
            if net_value == None:
                net_value = ''
        if column == "I":
            lpj_person = sheet[cell_name].value
            if lpj_person == None:
                lpj_person =''
        if column == "J":
            case_code = sheet[cell_name].value
            if case_code == None :
                case_code = ''
        if column == "L":
            user_note = sheet[cell_name].value
            if user_note == None :
                user_note = ''

    if main_key:
        tracker_dict[main_key] = [weight, net_value ,lpj_person,case_code,user_note]

print(tracker_dict)
for k , v in tracker_dict.items():
    print(v[0],v[1],v[2],v[3],v[4],k )
    con.execute('''update central_tracker set weight = "%s" , net_value = "%s" , LPJ_person = "%s" , casecode="%s" , user_note="%s" where article_number = "%s" '''%( str(v[0]),v[1],v[2],v[3],v[4],k ))
    con.commit()
