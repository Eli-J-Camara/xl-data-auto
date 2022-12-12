import random
import xlsxwriter

def main():

    # The Workbook function creates a new excel file.
    workbook = xlsxwriter.Workbook('apples.xlsx')

    bold_format = workbook.add_format({'bold': True})

    cell_format = workbook.add_format()
    cell_format.set_text_wrap()
    cell_format.set_align('top')
    cell_format.set_align('left')

    money_format = workbook.add_format({'num_format': '$#,##0.00'})
    money_format.set_align('top')
    money_format.set_align('left')

    money_red_format = workbook.add_format({'num_format': '$#,##0.00'})
    money_red_format.set_font_color('red')
    money_red_format.set_align('top')
    money_red_format.set_align('left')
    

    worksheet = workbook.add_worksheet('Apples')
    #This allows us to write to each individual cell.
    cells = ['A1', 'B1', 'C1', 'D1', 'E1']
    content = ['Tree ID', 'Number of Apples', 'Tree Height', 'Tree Profit', 'Tree Type']
    
    for i in range(5):
        worksheet.write(cells[i], content[i], bold_format)
    
    rowIndex = 2
    for row in range(200):
        treeid = row + 1000
        numApples = 20 + random.randint(50,100)
        typeOfTree = random.choice(['Macintosh', 'Red Delicious', 'Granny Smith', 'Fuji'])
        treeProfit = random.random() * 1000
        heightOfTree = 100 + random.randint(25,50)

        worksheet.write(f'A{rowIndex}', treeid, cell_format)
        worksheet.write(f'B{rowIndex}', numApples, cell_format)
        worksheet.write(f'C{rowIndex}', heightOfTree, cell_format)

        if treeProfit > 500:
            worksheet.write(f'D{rowIndex}', treeProfit, money_red_format)
        else:
            worksheet.write(f'D{rowIndex}', treeProfit, money_format)
        worksheet.write(f'E{rowIndex}', typeOfTree, cell_format)

        rowIndex += 1

        
    worksheet.set_column(4,4,width=20)
    # It is important to close the excel file you created, or it may not be written out.
    workbook.close()


if __name__ == "__main__":
    main()


