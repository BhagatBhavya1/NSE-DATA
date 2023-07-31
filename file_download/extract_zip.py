import zipfile
from read_csv import save_to_excel
# from read_csv import read_csv_data

def extract_zip(file_name):
    # Replace 'your_zip_file.zip' with the name of your zip file
    zip_file_path = f'C:/Users/HP/OneDrive/Documents/Downloads/{file_name}.zip'

    # Replace 'extracted_folder_path' with the path where you want to extract the contents of the zip file
    extracted_folder_path = 'Downloads'

    # Open the zip file
    with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
        # Extract all the contents of the zip file
        zip_ref.extractall(extracted_folder_path)

    print("Zip file extracted successfully!")
    save_to_excel(file_name)
    # read_csv_data(file_name)

