#https://github.com/demoyhui/Python-Application
#encoding:utf_8
#pip install xlrd #python 2.75
#only support python 2.75

__metaclass__ = type
import os
import glob
import xlrd
import sys
reload(sys)
sys.setdefaultencoding('utf-8')
path = os.getcwd()

def is_float_by_except(num):  
    try:  
        float(num)  
        return True  
    except ValueError:  
        #print "%s ValueError" % num  
        return False  
    return False
        

#print "path=%s" % path
#print "__file__=%s" % __file__
#print "os.path.realpath(__file__)=%s" % os.path.realpath(__file__)

file_input_path  = path + '/excel/*'
file_output_path = path + '/ini/'

files = glob.glob(file_input_path)

for file in files:
    wb = xlrd.open_workbook(file)

    for sheetName in wb.sheet_names():
        #print(sheetName)
        sheet = wb.sheet_by_name(sheetName)
        #print(sheet.nrows)
        # check duplicate keys
        l_key_list = [] 
        #sheet.nrows sheet.ncols
        for rownum in range(2,sheet.nrows):
            l_key = sheet.cell(rownum, 0).value
            l_key = str(l_key)
            l_key = l_key.rstrip()
                
            if l_key == "":
                continue
            if is_float_by_except(l_key):
                l_key = str(int(float(l_key)))
                    
            l_key_list.append(l_key)
        l_key_set = set(l_key_list)
        for item in l_key_set:
            if l_key_list.count(item) > 1:
                print("the %s has found %d" %(item, l_key_list.count(item)))
                #sys.exit(0)
            
        for colnum in range(1,18):
            lang_v = sheet.cell(1, colnum).value
            lang_v = str(lang_v)
            lang_v = lang_v.rstrip()            
            #print(lang_v)           
            lang_file_name = file_output_path + 'text_' + lang_v + '.ini'
            lang_file = open(lang_file_name, 'w')
            
            for rownum in range(2,sheet.nrows):
                #Excel treats all numbers as floats. 
                #In general, it doesn't care whether your_number % 1 == 0.0 is true or not.
                l_key = sheet.cell(rownum, 0).value
                l_key = str(l_key)
                l_key = l_key.rstrip()
                if l_key == "":
                    continue
                if is_float_by_except(l_key):
                    l_key = str(int(float(l_key)))
                    
                l_val = sheet.cell(rownum, colnum).value                
                l_val = str(l_val)
                l_val = l_val.rstrip()
                dataStr = l_key + '=' + l_val + '\n'
                lang_file.write(dataStr)
            
            lang_file.close()

print('success!!!')            
            
