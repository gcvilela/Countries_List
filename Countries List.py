import xlsxwriter
import requests

def Countries_list():
    #Make a request to the API, save it in a variable "r" and convert the file to json
    r = requests.get("https://restcountries.com/v3.1/all")
    all = r.json()

    #Create the worksheet and the Workbook tab
    workbook = xlsxwriter.Workbook('Countries List.xlsx')
    worksheet = workbook.add_worksheet()

    #Customization of title, subtitle and area cells
    sub = workbook.add_format({ 'bold': 1,'font_color': '#808080' ,'font_size': 12})
    title = workbook.add_format({ 'bold': 1,'font_color': '#4F4F4F' ,'font_size': 16,'align': 'center'})
    area_format = workbook.add_format({'num_format': ' #,##0.00'})

    #Place, center and merge the title
    worksheet.merge_range('A1:D1', 'Countries List', title)

    #Put the subtitles
    worksheet.write('A2', 'Name', sub)
    worksheet.write('B2', 'Capital', sub)
    worksheet.write('C2', 'Area', sub)
    worksheet.write('D2', 'Currencies', sub)

    sup_list = [] 
    sup_list2 = []

    for i in all:

        #Add it to the support list to be able to sort by putting each value in a variable
        name = (i["name"]["common"])

        try:
            cap = (i["capital"][0])
        except KeyError as ke:
            cap = "-"

        area = i["area"]

        #Add the value of all 'Currencies'
        g = ""
        try:
            for i in list(i["currencies"]):
                g = g + i +", "
            cur = g[:-2]
        except KeyError as ke:
            cur = "-"

        #CPut the list on another support list
        sup_list = [name , cap , area , cur]
        sup_list2.append(sup_list)

    #Organize in ascending order
    sup_list2.sort()

    #Write in xlsx for each row and column the saved value
    n = 2
    for i in sup_list2:
            worksheet.write(n , 0 , i[0])
            worksheet.write(n , 1 , i[1])
            worksheet.write(n , 2 , i[2] , area_format)
            worksheet.write(n , 3 , i[3])
            n+=1
    #close and show Finished message
    workbook.close()
    print("Finished")
    
    
Countries_list()
