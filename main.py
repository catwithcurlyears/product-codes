import pandas as pd
from texture_dict import texture_dict


# Path for output
text_file_path = r'C:\Users\june\Desktop\data.txt'

# Path for input and set up
excel_file_path = r'C:\Users\june\Desktop\data.xlsx'
sheet_name = 'Sheet1'
xls = pd.ExcelFile(excel_file_path)
column_name = 'EilutesPreke'
df = xls.parse(sheet_name)

if column_name in df.columns:
    data_values = df[column_name]

    texture_list = []
    electricity_values = ("-EU", "-UK", "-FR", "-CH", "-US")

    for i in range(0, len(data_values)):
        item = data_values[i]
        new_list = []
        code_number = ""
        texture_codes = ""
        # Get the code number and texture codes
        # Example: NZ65035G04-C60C60M01
        # code_number = NZ65035G04
        # texture_codes = C60C60M01

        if pd.notna(item) and (isinstance(item, str) or isinstance(item, int)):
            if item[0] == "N":
                code_number = item[:10]
                texture_codes = item.split('-')[1]
            elif item[0] == "G":
                if item[5].isalpha():
                    code_number = item[:5]
                    texture_codes = item[5:]
                else:    
                    code_number = item[:6]
                    texture_codes = item[6:]

            new_list.append(code_number)

            if len(texture_codes) > 2: 
                while len(texture_codes):
                    texture_check = texture_codes[:4]

                    # Check texture versions
                    if (texture_check[-1] in ["0", "1", "2", "3"]):
                        texture_codes = texture_codes[:3] + texture_codes[4:]
                    texture_check = texture_check[:3]
                    texture_codes = texture_codes[len(texture_check):]

                    # Check of electricity value
                    if any(i in texture_check for i in electricity_values):
                        texture_check = texture_check[1:]
                    
                    # If string contains " - " remove it
                    texture_check = texture_check.replace("-", "")

                    # Check dict and repalce valus they match
                    if texture_check in texture_dict:
                        texture_check = texture_dict[texture_check]

                    new_list.append(texture_check)

            # hecks if texture value is less than 3 andd print value
            elif (len(texture_codes) == 2) or (len(texture_codes) == 1):    
                new_list.append(texture_codes)

            texture_list.append(new_list)

    # Write textures to .txt file
    with open(text_file_path, "a") as text_file:
        for texture in texture_list:
            text_file.write('.'.join(texture) + '\n')

print('Done')

