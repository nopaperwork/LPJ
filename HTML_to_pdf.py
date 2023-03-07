from hashlib import new
from unittest.main import MAIN_EXAMPLES
from selenium import webdriver
import time
import os
import json
import re
from config import OUTPUT_HTML_FOLDER

from config import OUTPUT_PDF_FOLDER
from config import CHROME_PATH

# Now convert the html files to PDF
chrome_options = webdriver.ChromeOptions()
settings = {
            "recentDestinations": [{
                "id": "Save as PDF",
                "origin": "local",
                "account": "",
            }],
            "selectedDestinationId": "Save as PDF",
            "version": 2
            }
output_pdf_path = os.path.abspath(OUTPUT_PDF_FOLDER) 
prefs = {'printing.print_preview_sticky_settings.appState': json.dumps(settings),
            'savefile.default_directory': output_pdf_path}
            # 'savefile.default_directory':  OUTPUT_PDF_FOLDER}

file_list = os.listdir(OUTPUT_HTML_FOLDER)
chrome_options.add_experimental_option('prefs', prefs)
chrome_options.add_argument('--kiosk-printing')
driver = webdriver.Chrome(executable_path= CHROME_PATH,chrome_options=chrome_options)

for file in file_list :
    if file.endswith('.html'):
        output_file =os.path.abspath(OUTPUT_HTML_FOLDER + file )
        filename = 'file://' + output_file
        driver.get(filename)
        try :
            os.remove(OUTPUT_PDF_FOLDER+'Track Consignment.pdf')
        except :
            print('No old files found')
        # try:
        driver.execute_script('window.print();')
        # except FileExistsError:
        #     print('File already exist for :' , file)
            
        new_filename = OUTPUT_PDF_FOLDER+ file.replace('.html','.pdf')
        time.sleep(0.5)
        try:
            os.rename( OUTPUT_PDF_FOLDER+'Track Consignment.pdf', new_filename)
        except FileExistsError:
            print('File already exist for :' , new_filename)
            print('Updating with latest info')
            os.remove(new_filename)
            os.rename( OUTPUT_PDF_FOLDER+'Track Consignment.pdf', new_filename)
            
            