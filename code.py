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

print(master_df)
print(master_df.columns)


sf = StyleFrame(master_df)
sf.apply_style_by_indexes(indexes_to_style=sf[sf['note'] < 3],
                          cols_to_style=['Nom'],
                          styler_obj=Styler(bg_color='green', bold=True, font_size=10))

sf.to_excel('master_df.xlsx').save()