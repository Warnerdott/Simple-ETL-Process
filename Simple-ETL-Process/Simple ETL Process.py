import pyodbc 
import pandas 
import pandasql 

#SETUP THE MS SQL
Driver = "DRIVER={ODBC Driver 17 for SQL Server};"
Server = "SERVER=IS-SOFTDEV9-NB\\SQLEXPRESS;"
Database = "DATABASE=Practice_DB;" 
connectionString = pyodbc.connect(Driver + Server + Database + "Trusted_Connection=yes;")

#LOAD THE EXCEL FILE
sample_excel = pandas.ExcelFile(r'C:\Users\1023242\Desktop\DE\SampleData.xlsx')
sheets = sample_excel.sheet_names

#Checking the SalesOrders in sheets
if 'SalesOrders' in sheets:
        SalesOrders_Table = pandas.read_excel(sample_excel, sheet_name='SalesOrders')
        SalesOrders_Table['OrderDate'] = SalesOrders_Table['OrderDate'].dt.strftime('%Y-%m-%d')

        #TRANSFORM
        query = "SELECT * FROM SalesOrders_Table WHERE Region = 'Central' AND Units > 50 "
        excel_result = pandasql.sqldf(query)  
        #print(result.to_string (index=False))

        if not excel_result.empty:
            connection = connectionString.cursor()
            #CHECKING THE TABLE IF EXIST
            check_table_query = "IF OBJECT_ID('dbo.SalesOrders', 'U') IS NOT NULL SELECT 1 ELSE SELECT 0"
            table_exists = connection.execute(check_table_query).fetchone()[0]

            #DROP THE TABLE
            if table_exists:
                drop_query = "DROP TABLE dbo.SalesOrders"
                connection.execute(drop_query)
                connection.commit()
                print("DROP TABLE")

            #CREATE TABLE 
            create_query = "CREATE TABLE dbo.SalesOrders (OrderDate DATE, Region VARCHAR(255), Rep VARCHAR(255), " \
            " Item VARCHAR(255), Units INT, Unit_Cost FLOAT, Total FLOAT) "
            connection.execute(create_query)
            connection.commit()
            print("SalesOrders table is created")

                #LOAD THE DATA TO THE TALBE
            for index, row in excel_result.iterrows():
                insert_query = "INSERT INTO dbo.SalesOrders (OrderDate, Region, Rep, Item, Units, Unit_Cost, Total) " \
                f"VALUES ('{row['OrderDate']}', '{row['Region']}', '{row['Rep']}', '{row['Item']}', " \
                f"{row['Units']}, {row['Unit Cost']}, {row['Total']})"
                connection.execute(insert_query)
                connection.commit()
            print("SUCESS INSERT")

        else:
            print("No output from SalesOrders sheet.")
else:
     print("There is no SalesOrders sheet in excel.")