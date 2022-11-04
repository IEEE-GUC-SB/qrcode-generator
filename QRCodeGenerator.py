# Importing libraries
import qrcode, openpyxl, sys 

# Ensure correct usage
if len(sys.argv) != 2:
    sys.exit("Usage: python3 QRCodeGenerator.py <Excel file>")
fileName = sys.argv[1]

# Define variable to load the dataframe
dataframe = openpyxl.load_workbook(fileName)

# Define variable to read sheet
dataframe1 = dataframe.active

dataframe1.cell(1, dataframe1.max_column + 1).value = "ID"

dataframe1.cell(1, dataframe1.max_column + 1).value = "QRCode"

# Iterate the loop to read the cell values
for row in range(1, dataframe1.max_row):
    # Data to be encoded
    data = ''
    dataframe1.cell(row + 1, dataframe1.max_column-1).value = row
    for col in dataframe1.iter_cols(1, dataframe1.max_column):
        data += str(col[row].value) + ' '
    # Encoding data
    img = qrcode.make(data)
    # Saving as an image file
    img.save('MyQRCode' + str(row) + '.png')
    dataframe1.cell(row + 1, dataframe1.max_column ).hyperlink = 'C:\\Users\\hp\\Documents\\software\\SE-task1-team5\\MyQRCode' + str(row) +'.png'
    
dataframe.save(fileName)
