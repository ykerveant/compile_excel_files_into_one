#  https://stackoverflow.com/questions/54109548/how-to-save-pandas-to-excel-with-different-colors

import os
import pandas as pd
from styleframe import StyleFrame, Styler

def create_list_of_df_from_xlsx_files():
    directory = r"C:/Users/ykerv/Documents/classe emilie"
    file_list = [filename for filename in os.listdir(directory)]
    upgraded_file_list = [pd.read_excel(directory+"/"+file) for file in file_list]
    return upgraded_file_list


print(create_list_of_df_from_xlsx_files())
print("*|"*50)
# upgraded_file_list = [pd.read_excel("C:/Users/ykerv/Documents/classe emilie/"+file) for file in file_walk()]

print(create_list_of_df_from_xlsx_files())

#df = pd.read_excel("C:/Users/ykerv/Documents/classe emilie/Copie de 102 - Émilie D.xlsx")
df0 = pd.read_excel("file1.xlsx")

print(df0)
print(df0.columns)

df1 = pd.read_excel("file2.xlsx")

print(df1)
print(df1.columns)

frames = [df0, df1]
# when in production, replace frames for create_list... function
master_df = pd.concat(frames, ignore_index=True)
master_df.sort_values(by='Id', inplace=True)

# end of master_df
print(master_df)
print(master_df.columns)

# creating sf dataFrame for color formating
sf = StyleFrame(master_df)
cols_to_style = [cols for cols in sf.columns]

# applying color for student with 1
sf.apply_style_by_indexes(indexes_to_style=sf[(sf['note'] <= 1)],
                          cols_to_style=cols_to_style,
                          styler_obj=Styler(bg_color='green', bold=True, font_size=10))

# applying color for student with 2
sf.apply_style_by_indexes(indexes_to_style=sf[(sf['note'] > 1) & (sf['note'] < 3)],
                          cols_to_style=cols_to_style,
                          styler_obj=Styler(bg_color='yellow', bold=True, font_size=10))

# applying color for student with 3
sf.apply_style_by_indexes(indexes_to_style=sf[(sf['note'] >= 3)],
                          cols_to_style=cols_to_style,
                          styler_obj=Styler(bg_color='red', bold=True, font_size=10))

sf.to_excel('master_df.xlsx').save()
