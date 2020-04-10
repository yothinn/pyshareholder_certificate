from PIL import Image, ImageDraw, ImageFont
from oauth2client.service_account import ServiceAccountCredentials
from pythainlp.util import bahttext, thai_strftime
from datetime import datetime
import gspread

# Transaction sharesholder google sheet and sheet name
GGSHEET_ID = "1blvVgkPMWyC5x_B2c7gx0NfozQsxVSGR_Un1zh_aWMg"
GGSHEET_NAME = "transactions"
CRED_FILE = "./env/thamturakit-data-center-credential.json"
OUTPUT_PATH = "./output/"
FONT_FILE = "./env/THSarabunIT.ttf"
FONT_SIZE = 32
CERT_POS = {
    "shareholder_id": (100, 100),
    "no_cert": (100, 100),
    "prefix": (100, 100),
    "firstname": (100, 100),
    "lastname": (100, 100),
    "num_share": (100, 100),
    "share_amount": (100, 100),
    "share_bathtext": (100, 100),
    "cert_date": (100, 100)
}

'''
convert google sheet row to shareholder data
@param {array of string} gg_row :  data from google sheet in one row
return shareholder data object
'''
def toshareholderdict(ggdata):
    return {
        "shareholder_id": ggdata[0],
        "no_cert": ggdata[12],
        "prefix": ggdata[17],
        "firstname": ggdata[18],
        "lastname": ggdata[19],
        "num_share": ggdata[10],
        "share_amount": ggdata[8],
        "cert_date": ggdata[3] 
    }


'''
create shareholder certificate jpeg image
@param {dict} data dict
@param {string} filename to save jpeg image
return true for success create jpeg image file
'''
def shareholder_cert(datadict, filename):
    print(CERT_POS)
    print(datadict)

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

    # font = ImageFont.truetype('THSarabunIT.ttf', size=32)

    while(True):

        # Input part
        str = input('Input your row (or "exit") : ')

        if str=='exit':
            exit()
        elif not str.isdecimal():
            print('!!Oop:: input Error please input row')
            continue

        # Get data from google sheet
        try:
            row_data = worksheet.row_values(int(str))           
            datadict = toshareholderdict(row_data)
            filename = "{}{}".format(OUTPUT_PATH, "test.jpg")
            print(row_data)
            # print(datadict)
            shareholder_cert(datadict, filename)
        except:
            print ('Cant load data in {} row'.format(str))
            continue
   
    # # col 3 = first name
    # # col 4 = last name
    # if row_list[3] == "":
    #     print ('May be Empty data, please check')
    #     continue

    # full_name = '{}  {}'.format(row_list[3].strip(), row_list[4].strip())

    # print(row_list)
    # print(full_name)

    # # ข้อมูลลงขันแบบมีส่วนร่วม
    # # col 16 = participate amount
    # if not row_list[16] == "":
    #     participate_amount = float(row_list[16].strip().replace(',', ''))
    #     if not participate_amount == 0:
    #         participate_date = datetime.strptime(row_list[25], '%d/%m/%Y')

    #         participate_image = Image.open("participate_certificate_bg.jpg")
    #         participate_draw = ImageDraw.Draw(participate_image)

    #         participate_draw.text((220,260), '{}  '.format(full_name), "rgb(0, 0, 0)", font=font)
    #         participate_draw.text((220, 320), '{:,.2f}'.format(participate_amount), "rgb(0, 0, 0)", font=font)
    #         participate_draw.text((600, 320), '( {} )'.format(bahttext(participate_amount)), fill="rgb(0, 0, 0)", font=font)
    #         participate_draw.text((220, 385),  thai_strftime(participate_date, '%d  %B  %Y'), fill="rgb(0, 0, 0)", font=font)

    #         participate_filename = '{}_ส่วนร่วม_{}.jpg'.format(row_list[31].strip().replace("/", "_"), full_name)
    #         participate_image.save(participate_filename)

    #         # clear image
    #         del participate_draw
    #         participate_image.close()

    # # ข้อมูลลงขันแบบผลตอบแทน         
    # if not row_list[15] == "":
    #     invest_amount = float(row_list[15].strip().replace(',', ''))
    #     print('({})'.format(bahttext(1000000.0))) 

    #     if not invest_amount == 0:
    #         invest_date = datetime.strptime(row_list[24], '%d/%m/%Y')

    #         invest_image = Image.open("invest_certificate_bg.jpg")
    #         invest_draw = ImageDraw.Draw(invest_image)

    #         invest_draw.text((220,260), '{}  '.format(full_name), "rgb(0, 0, 0)", font=font)
    #         invest_draw.text((220, 320), '{:,.2f}'.format(invest_amount), "rgb(0, 0, 0)", font=font)
    #         invest_draw.text((600, 320), '( {} )'.format(bahttext(invest_amount)), fill="rgb(0, 0, 0)", font=font)
    #         invest_draw.text((220, 385),  thai_strftime(invest_date, '%d  %B  %Y'), fill="rgb(0, 0, 0)", font=font)

    #         invest_filename = '{}_ผลตอบแทน_{}.jpg'.format(row_list[29].strip().replace("/", "_"), full_name)
    #         invest_image.save(invest_filename)

    #         # clear image
    #         del invest_draw
    #         invest_image.close()
    
    
if __name__ == "__main__":
    main()
