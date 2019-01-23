# -*- coding: utf-8 -*-
"""
Created on Tue Nov 20 12:45:03 2018

@author: gaoyu
"""
'''
Step2: Query from SQL to copy the whole database to a local excel SUTD.xlsx with sheets named after table names in SQL
so people who have the access to the SQL database and download it in a excel format
'''
import pyodbc
import pandas as pd

#copySQLtoExcel(tablenames,excelname), tablenames is a list of names of tables in excel to be downloaded. the data downloaded will be transformed to a excel named after the excelname
def copySQLtoExcel(tablenames,excelname):
    #create a excel to write the data downloaded to, if name existing, will write to that existing excel
    writer = pd.ExcelWriter(excelname+'.xlsx')
    # for all tables with names in the tablenames list, will be downloaded
    for tablename in tablenames:
        #making connection with SQL database
        #below is a sample for remote server
        # #connection=pyodbc.connect(r'Driver={SQL Server};SERVER=exactcurefirstdb.cb9zgxqtumuv.eu-west-1.rds.amazonaws.com;DATABASE=SUTDstudent;UID=mbexactcure;PWD=ExactCurePower42')
        connection = pyodbc.connect(r'Driver={SQL Server};Server=localhost;Database=SUTDstudent;Trusted_Connection=yes;')
        script ="SELECT * FROM "+tablename   
        #get the pandas dataframe from SQL command 
        df= pd.read_sql_query(script, connection) 
        # write the dataframe to excel sheet named after its original table name   
        df.to_excel(writer, tablename)
    writer.save()
    #save the changes in the excel

def main():
    #edit the tablenames list to choose the tables in SQL database that you want to download
    tablenames=["Year10","Year11","Year12"]
    #create or choose the excel file to save the downloads
    excelname="SUTD"
    #call the function
    copySQLtoExcel(tablenames,excelname)

if __name__ == '__main__':
    main()
    