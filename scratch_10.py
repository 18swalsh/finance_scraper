import pandas as pd

test_dict = {}


df = pd.DataFrame({'numbers':    [1010, 2020, 3030, 2020, 1515, 3030, 4545],
                   'percentages': [.1,   .2,   .33,  .25,  .5,   .75,  .45 ],
})

df_two = pd.DataFrame({'mixed':    ["QWE", [123,'abc'], 3030],
                   'lists': [[1,2,3],   ['a','b','c'],  ["one_item"]],
})

letters = 'ABCDEFGHI'

for tick in range(0,2):
    test_dict[letters[tick]+letters[tick]+letters[tick]+letters[tick]]="test_string"

for tick in range(3,5):
    test_dict[letters[tick]+letters[tick]+letters[tick]+letters[tick]]= ["string","string_two","string_three",(1,3),df]

for tick in range(6,7):
    test_dict[letters[tick]+letters[tick]+letters[tick]+letters[tick]]="test_string"

for tick in range(8,9):
    test_dict[letters[tick]+letters[tick]+letters[tick]+letters[tick]]= df_two

print test_dict

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter("pandas_column_formats.xlsx", engine='xlsxwriter')
x = 0
# Convert the dataframe to an XlsxWriter Excel object.
for key in test_dict:
    x += 1
    test_dict[key].to_excel(writer, sheet_name = 'Sheet' + str(x))
    #           df.to_excel(writer, sheet_name='Sheet1')

# Get the xlsxwriter workbook and worksheet objects.
#workbook  = writer.book
#worksheet = writer.sheets['Sheet1']

# Add some cell formats.
#format1 = workbook.add_format({'num_format': '#,##0.00'})
#format2 = workbook.add_format({'num_format': '0%'})

# Note: It isn't possible to format any cells that already have a format such
# as the index or headers or any cells that contain dates or datetimes.

# Set the column width and format.
#worksheet.set_column('B:B', 18, format1)

# Set the format but not the column width.
#worksheet.set_column('C:C', None, format2)

# Close the Pandas Excel writer and output the Excel file.
writer.save()