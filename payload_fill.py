from seleniumwire import webdriver  # Import from seleniumwire
from config import CHROME_PATH
# Create a new instance of the Chrome driver
driver = webdriver.Chrome(executable_path= CHROME_PATH)

# Go to the Google home page
driver.get('https://www.indiapost.gov.in/_layouts/15/DOP.Portal.Tracking/TrackConsignment.aspx')

input('Is the Response visible')
# Access requests via the `requests` attribute
for request in driver.requests:
    if request.url == 'https://www.indiapost.gov.in/_layouts/15/DOP.Portal.Tracking/TrackConsignment.aspx':
        if request.response and request.response.status_code == 200:
            print(
                request.body,
            )
            final_word = str(request.body)
            print('________________________________________________________________________')

try:
    with open('payload.txt','w') as f:
        f.write(final_word)
    print('Successfully updated Payload file')
except:
    print('Something got failed')
    
input()