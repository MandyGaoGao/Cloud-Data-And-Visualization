# -*- coding: utf-8 -*-
"""
Created on Tue Jan 22 21:19:51 2019

@author: gaoyu
"""
'''
Data.xlsx Excel: Original local database
Step1: Create SQL database and populate data from Data.xlsx to the cloud SQL server e.g "SUTDstudent" database
'''
import pandas as pd #lib that handles data
import pyodbc#library that handles Microsoft SQL server with python
#################### Create a table in SQL server data base, table structure==excel structure ##################################
def createSQLtable(file,tablename,connection):
    #read the sheet named by tablename from the input excel
    a=pd.read_excel('{}.xlsx'.format(file),tablename)
    #get column names of the excel
    colnames=a.columns.values
    #create SQL command string based on the tablename and column names
    #this command is to create table
    command=""
    for i in colnames:  
        command+=i+" nvarchar(100)"
        if i==colnames[-1]:
            command+=")"
        else:
            command+=","
    SQLCommand0="CREATE TABLE "+tablename+" ("+command
    #make the input connection with SQL and commit the command
    cursor = connection.cursor()
    cursor.execute(SQLCommand0)
    connection.commit() 
    connection.close()
####################### upload the excel info to the SQL table created just now ##############################
def populateSQLfromExcel(file,tablename,connection):
    #read the input excel and load all content to a list called lst1
    a=pd.read_excel('{}.xlsx'.format(file),tablename)    
    names=a.columns
    colnames=a.columns.values #get column names of the excel
    lst1=[]
    for i in a.index:
        lst=[]
        for name in names:
            lst.append(str(a[str(name)][i]))
        lst1.append(lst)       
    #create SQL command string based on the tablename and column names
    #this command is to update info
    command2="INSERT INTO dbo."+tablename+" ("
    for i in colnames:
        command2+=i
        if i==colnames[-1]:
            command2+=") VALUES ("
        else:
            command2+=","
    for i in colnames:
        command2+="?"
        if i==colnames[-1]:
            command2+=")"
        else:
            command2+="," 
    SQLCommand = command2  
    #set up input SQL server connection and commit SQL command
    cursor2 = connection.cursor() 
    for i in lst1:  
        Values =i
        cursor2.execute(SQLCommand,Values) 
    connection.commit() 
    connection.close()#close the connection after done
#################### main process #################################
def main():
    #type in the name of the input excel here as a string
    file="Data"
    #type in the sheet name of the input excel that you want to update to the database as strings in a list
    tablenames=["Year10","Year11","Year12"]
    # for all excel sheets to update to SQL, createSQLtable,populateSQLfromExcel will be done 
    for tablename in tablenames:
        connection = pyodbc.connect(r'Driver={SQL Server};Server=localhost;Database=SUTDstudent;Trusted_Connection=yes;')
        createSQLtable(file,tablename,connection)
        connection = pyodbc.connect(r'Driver={SQL Server};Server=localhost;Database=SUTDstudent;Trusted_Connection=yes;')
        populateSQLfromExcel(file,tablename,connection)
    #here connection is local host
    #the following is a sample of the potential code for remote sever,here uses aws
    #connection=pyodbc.connect(r'Driver={SQL Server};SERVER=exactcurefirstdb.cb9zgxqtumuv.eu-west-1.rds.amazonaws.com;DATABASE=SUTDstudent;UID=mbexactcure;PWD=ExactCurePower42')
#####################################################       
if __name__ == '__main__':
    main()