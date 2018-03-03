
#we're going to make a function that converts the columnNumber to a letter since in excel it's a1 not 1,1 
def rowColConvert(row, col):
    #if the row is greater than 26 (after z) it adds max one set of letters to it.
    if int(col)/26 > 0:
        colLetter = chr(int(col)/26+65) + chr(int(col)%26+65)
    else:
        #otherwise; we just convert it
        colLetter = chr(65+col)
    #python is semi hard typed; so like you can't just concatinate stuff in python like you can in javascript; still easier than c++ tho. 
    return str(colLetter)+str(row)
