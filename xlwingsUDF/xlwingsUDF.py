
# Start a sheet project with:  xlwings quicstart --standalone <project> and this will include xlwings as a VBA module instead of using the addin.
# In order to use xlwings without the addin, you can edit the _xlwings.conf sheet in the workbook and activate it by renaming it xlwings.conf (drop the "_")
# xlwings.conf filepath for interpreter must have path to pythonw.exe.   Note sometimes the xlwings*.dll may be in the wrong location.  They should be in the 
# same folder as python.exe and pythonw.exe

import xlwings as xw
import datetime as dt


@xw.sub  # only required if you want to import it or run it via UDF Server
def main():
    wb = xw.Book.caller()
    wb.sheets[0].range("A1").value = "Hello xlwings!"


@xw.func
def hello(name):
    return "hello {0}".format(name)
    
@xw.func
def myfunc(myarg):
    # check myarg for specific type
    if type(myarg)==dt.datetime:
        return "Return the date {}/{}/{}!!!!!".format(myarg.day, myarg.month, myarg.year)
    else:
        return "Return the variable {}".format(myarg)


if __name__ == "__main__":
    xw.books.active.set_mock_caller()
    
    main()          #<-When not debugging, uncomment this line and set "Use UDF Server" and "Debug UDFs" to FALSE
    #xw.serve()     #<-This is needed to debug the Python code using the UDF server.   Make sure that xlwings.conf '"Use UDF server" and "UDF Debug" set to TRUE