import xlrd
#from xlrd import cellname #this is not needed unless you want to use the name of cell e.g A1

def open_xlBook(file_name, sheet_name):
    data = xlrd.open_workbook(file_name)
    sheet = data.sheet_by_name(sheet_name)
    return sheet

def print_contents(sheet):
    for row_index in range(sheet.nrows):
        for col_index in range(sheet.ncols):
            #cell_name = cellname(row_index,col_index)    
            print sheet.cell(row_index,col_index).value

def createDictionary(sheet):
    name_sal_dict = {} #create a one to one mapping between names and salaries
    sal = [] #create list to store salaries
    for row_index in range(sheet.nrows):
        for col_index in range(sheet.ncols):
            #skip the title row (i.e topmost row) and only focus on column 0
            #where we get all the names
            if row_index > 0 and col_index <1:
                #get the salary from 3rd column and map it to the name from column 0
                name_sal_dict[sheet.cell(row_index,col_index-col_index).value] = \
                                            sheet.cell(row_index,col_index+2).value
                #keep adding the salaries to list
                sal.append(sheet.cell(row_index,col_index+2).value)   
            else:
                pass
    return name_sal_dict,sal

def whoEarnMost(dic, lis):
    #assemble the result string to write to the file
    output_string = ""
    for name, sal in dic.iteritems():
        if sal == max(lis):
            output_string = output_string + name + " earns " + str(sal) + " USD."+'\n'
    return output_string
    
    
def writeOut(output_string):
    f = open('output.txt','w')
    f.write(output_string)
    f.close
    
if __name__=='__main__':
    xlFile = './Information.xlsx'
    sheetName = 'People'
    #open sheet
    sheet = open_xlBook(xlFile, sheetName)
    #get the name: salary dictionary and the salary list
    name_salary_dict, salary_list = createDictionary(sheet)
    #get the result string
    out_string = whoEarnMost(name_salary_dict, salary_list)
    #output to the file
    writeOut(out_string)
