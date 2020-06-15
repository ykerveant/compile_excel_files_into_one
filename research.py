"""
# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('master.xlsx', engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.
master_df.to_excel(writer, sheet_name='master_df', index=False)

# Get the xlsxwriter workbook and worksheet objects.
workbook = writer.book
worksheet = writer.sheets['master_df']

# Close the Pandas Excel writer and output the Excel file.
writer.save()
"""

"""
####################
sf = StyleFrame({'a': ['a', 'b', 'c', 'd']})
yellow = Styler(bg_color='yellow')
blue = Styler(bg_color='blue')

sf.apply_style_by_indexes(sf.index[0], yellow)
sf.apply_style_by_indexes(sf.index[1], blue)
sf.to_excel().save()
##################
"""