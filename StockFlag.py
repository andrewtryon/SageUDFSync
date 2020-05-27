import string
import pyodbc
import pandas as pd
import subprocess
import csv
import openpyxl
import requests
import os
import sys
import shutil
from PIL import Image 
import datetime
from datetime import date

if __name__ == '__main__':

    #Connection String
    conn_str = (
        r'DSN=SOTAMAS90;'
        r'UID=********;'
        r'PWD=********;'
        r'Directory=\\sage100\sage 100 Advanced\2018\MAS90;'
        r'Prefix=\MAS90\SY\, \\sage100\Sage 100 Advanced\2018\MAS90\==\;'
        r'ViewDLL=\\sage100\Sage 100 Advanced\2018\MAS90\HOME;'
        r'Company=*****;'
        r'LogFile=\PVXODBC.LOG;'
        r'CacheSize=4;'
        r'DirtyReads=1;'
        r'BurstMode=1;'
        r'StripTrailingSpaces=1;'
        r'SERVER=NotTheServer;'
        )

    logf = open("download.log", "w")
    try:
        # code to process download here

        print('Connecting to Sage')
        cnxn = pyodbc.connect(conn_str, autocommit=True)    
        
        #SQL Sage data into dataframe
        sql = """
            SELECT 
                CI_Item.ItemCode
            FROM 
                CI_Item CI_Item, 
                IM_ItemWarehouse IM_ItemWarehouse
            WHERE 
                CI_Item.ItemCode = IM_ItemWarehouse.ItemCode AND
                IM_ItemWarehouse.WarehouseCode = '000' AND
                IM_ItemWarehouse.ReorderPointQty = 0 AND
                (CI_Item.UDF_NONSTOCK = 'N' or CI_Item.UDF_NONSTOCK is Null)     
        """

        PutOnNonStock = pd.read_sql(sql,cnxn) 
        print('AA_PUTONNONSTOCK_VIWI5Q')
        if PutOnNonStock.shape[0] > 0:
            filepath = r'\\*******\Alt Team\Qarl\Automatic VI Jobs\Maintenance\CSVs\AA_PUTONNONSTOCK_VIWI5Q.csv' 
            PutOnNonStock.to_csv(filepath, index=False, header=False)
            print(PutOnNonStock)
            
            #Auto VI .... uncomment  below to turn on....untested
            p = subprocess.Popen('Auto_PutOnNonStock_VIWI5Q.bat', cwd=r"Y:\Qarl\Automatic VI Jobs\Maintenance", shell=True)
            stdout, stderr = p.communicate()
            print('Sage VI Complete!')
        else:
            print('No PutOnNonStock')     
            

        sql = """
            SELECT 
                CI_Item.ItemCode
            FROM 
                CI_Item CI_Item, 
                IM_ItemWarehouse IM_ItemWarehouse
            WHERE
                CI_Item.ItemCode = IM_ItemWarehouse.ItemCode AND
                IM_ItemWarehouse.WarehouseCode = '000' AND
                IM_ItemWarehouse.ReorderPointQty > 0 AND
                CI_Item.UDF_NONSTOCK = 'Y'
        """
        print('AA_TAKEOFFNONSTOCK_VIWI5P')
        TakeOffFlag = pd.read_sql(sql,cnxn)
        if TakeOffFlag.shape[0] > 0:
            filepath = r'\\FOT00WEB\Alt Team\Qarl\Automatic VI Jobs\Maintenance\CSVs\AA_TAKEOFFNONSTOCK_VIWI5P.csv' 
            TakeOffFlag.to_csv(filepath, index=False, header=False)
            print(TakeOffFlag)
            
            #Auto VI .... uncomment  below to turn on....untested
            p = subprocess.Popen('Auto_TakeOffNonStock_VIWI5P.bat', cwd=r"Y:\Qarl\Automatic VI Jobs\Maintenance", shell=True)
            stdout, stderr = p.communicate()
            print('Sage VI Complete!')
        else:
            print('No TakeOffFlag')                  
            
    except Exception as e:     # most generic exception you can catch
        logf.write("Error :( {0}\n".format(str(e)))
        # optional: delete local version of failed download
    finally:
        # optional clean up code
        pass
    #Establish sage connection