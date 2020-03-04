import xlsxwriter
import sqlite3

# Open the workbook and define the worksheet
book = xlsxwriter.Workbook("Quotedb.xlsx")
sheet = book.add_worksheet("dataimported")

# Establish a Sqlite connection
path = 'Quotedatabase.db'
db = sqlite3.connect(path)


# Get the cursor, which is used to traverse the database, line by line

cursor = db.cursor()

#Read data from the customerdb database file

cursor.execute('''select * from quotedb''')
all_rows = cursor.fetchall()

#initialize rows and columns of the worksheet
row = 0
col = 0

#Insert the columns name in to the excel
column_Values = [ 'Name' ,'Project_name','Filament_length(Hours)', 'Print_time(Hours)','Raw_material_cost_per_meter',
                                                        'Raw_material_cost',
                                                        'Power_consumption_cost',
                                                        'Machine_depreciation_cost',
                                                        'Total_mfg_cost',
                                                        'Number_of_grids_used',
                                                        'Number_of_hours_of_Post_process',
                                                        'Wet_sanding_cost',
                                                        'Total_post_process_cost',
                                                        'Total_design_cost',
                                                        'Total_slicing_cost',
                                                        'Total_shipping_cost',
                                                        'Total_Packaging_cost',
                                                        'Total_profit_cost',
                                                        'Internet_charges',
                                                        'Conversation_charges',
                                                        'Laptop_electricity_charges',
                                                        'Laptop_depreciation_charges',
                                                        'Admin_and_Marketing_costs',
                                                        'Rent_cost',
                                                        'Total_Misc_costs',
                                                        'Total_Project_Cost']
                                                        
                                                       
                                                        

for heading in column_Values:
      sheet.write(row,col,heading)
      col+=1
# Create a For loop to iterate through each entries in the db file
for entry in all_rows:
      row += 1
      col = 0
      for data_val in entry:
            sheet.write(row,col,data_val)
            col += 1

#Close the workbook
book.close()

# Close the cursor
cursor.close()

# Commit the transaction
db.commit()

# Close the database connection
db.close()

