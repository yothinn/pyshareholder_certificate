from oauth2client.service_account import ServiceAccountCredentials
from PIL import Image, ImageDraw, ImageFont
from pythainlp.util import bahttext, thai_strftime
from datetime import datetime
import gspread
import re
from shareholder import ShareholderData


# Transaction sharesholder google sheet and sheet name
GGSHEET_ID = "1blvVgkPMWyC5x_B2c7gx0NfozQsxVSGR_Un1zh_aWMg"
GGSHEET_NAME = "transactions"
CRED_FILE = "./config/thamturakit-data-center-credential.json"
OUTPUT_PATH = "./output/"

def main():
    # connect google Sheet api
    try:
        scope = ['https://www.googleapis.com/auth/spreadsheets.readonly']
        credential = ServiceAccountCredentials.from_json_keyfile_name(CRED_FILE, scope)
        # credential = ServiceAccountCredentials.from_json_keyfile_dict(cred, scope)
        client = gspread.authorize(credential)
    except: 
        print('Cant connect google sheet API')
        exit()

    # connect to google sheet
    try:    
        sheet = client.open_by_key(GGSHEET_ID)
        # RAW DATA
        worksheet = sheet.worksheet(GGSHEET_NAME)
    except:
       print('Cant coonnet google sheet')
       exit()

    print('*** Generate Participate/Investment Certificate ***')
    print('*** when exit  please type "exit" ')
 
    while(True):

        # Input part
        fr = input("Input your First row (or exit) : ")
        er = input("Input your End row: ")

        if fr == "exit" or er == "exit":
            exit()
        elif not fr.isdecimal() or not er.isdecimal():
            print('!!Oop:: input Error please input row')
            continue

        # Get data from google sheet
        try:
            for row in range(int(fr), int(er)+1):
                row_data = worksheet.row_values(row)
                
                # is null or not
                if (not row_data):
                    print("Empty data")
                    continue
                       
                data = ShareholderData(row_data)

                # TODO : check output folder ?

                # create filename with no certificate and firstname/lastname
                firstname = re.sub("[/\:*?<>|.]", "", data.firstname)
                lastname = re.sub("[/\:*?<>|.]", "", data.lastname)
                no_cert = data.no_cert.replace("/", "_")
                filename = "{}{}_{} {}.jpg".format(OUTPUT_PATH, no_cert, firstname, lastname)

                # Create certificate image
                status = data.create_cert(filename)
                if status:
                    print("Success create certificate in row {} to file {}".format(row, filename))
                else:
                    print("!!--Opp cant create certificate --!!")

                # print(row_data)
                # print("{} {} {}".format(data.title, data.firstname, data.lastname))
                # print("{} {} {}".format(data.shareholder_id, data.no_cert, data.cert_date))
                # print("{} {}".format(data.share_amount, data.num_share))
            
        except:
            print ('Cant load data in {} row'.format(str))
            continue  
    
if __name__ == "__main__":
    main()
