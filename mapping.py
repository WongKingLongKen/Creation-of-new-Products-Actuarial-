# mapping.py
# purpose: map between the original plancode and the new plancode, e.g. AAA -> AAB
import pandas as pd
import os 

working_path = r'C:\2024-09 (v2)\7. Table Working\TABLE_conversion\testing'

product_list = pd.read_csv(working_path + r'\product_list.csv')

output = '[\n'

for i in range(product_list.shape[0]):
    product_pair = product_list.loc[i, :].to_list()
    output = output + '(\'' + str(product_pair[0]) + '\',\'' + str(product_pair[1])  +'\')' + ',\n'
    print(product_pair)
    print(output)

with open(working_path + r'\text_to_py.txt', 'w') as file:
    file.write(output[:-1] + '\n]')

