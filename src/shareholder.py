from PIL import Image, ImageDraw, ImageFont
from pythainlp.util import bahttext, thai_strftime
from datetime import datetime

BGIMG_FILE = "./config/shareholder_masterform.jpg"
FONT_FILE = "./config/THSarabunIT.ttf"
FONT_SIZE = 72
FONT_COLOR = "rgb(0, 0, 0)"
# Certificate image position in pixel
CERT_POS = {
    "shareholder_id": (540, 630),
    "no_cert": (1560, 630),
    "name": (650, 760),
    "num_share": (1000, 900),
    "share_amount": (1420, 1045),
    "share_bahttext": (200, 1180),
    "cert_date": (550, 1310),
    "refer": (200, 1930)
}

# Google sheet index
GGSHEET_INDEX = {
    "shareholder_id": 0,
    "no_cert": 12,
    "title": 22,
    "firstname": 23,
    "lastname": 24,
    "num_share": 10,
    "share_amount": 8,
    "cert_date": 3,
    "line_name": 25,
    "fb_name": 26,
}


class ShareholderData:
    def __init__(self, ggsheet_row):
        l = len(ggsheet_row)
        self.shareholder_id = ggsheet_row[GGSHEET_INDEX["shareholder_id"]]
        self.no_cert = ggsheet_row[GGSHEET_INDEX["no_cert"]]
        self.title = ggsheet_row[GGSHEET_INDEX["title"]]
        self.firstname = ggsheet_row[GGSHEET_INDEX["firstname"]]
        self.lastname = ggsheet_row[GGSHEET_INDEX["lastname"]]
        self.num_share = ggsheet_row[GGSHEET_INDEX["num_share"]]
        self.share_amount = ggsheet_row[GGSHEET_INDEX["share_amount"]]
        self.cert_date = ggsheet_row[GGSHEET_INDEX["cert_date"]]
        if (GGSHEET_INDEX["line_name"] < l):
            self.line_name = ggsheet_row[GGSHEET_INDEX["line_name"]]
        else:
            self.line_name = ""

        #print("pass")
        if (GGSHEET_INDEX["fb_name"] < l):
            self.fb_name = ggsheet_row[GGSHEET_INDEX["fb_name"]]
        else:
            self.fb_name = ""


    @property
    def shareholder_id(self):
        return self.__shareholder_id

    @shareholder_id.setter
    def shareholder_id(self, var):
        self.__shareholder_id = ""
        if var:
            self.__shareholder_id = self.cleantext(var)          

    @property
    def no_cert(self):
        return self.__no_cert
    
    @no_cert.setter
    def no_cert(self, var):
        self.__no_cert = ""
        if var:
            self.__no_cert = self.cleantext(var)

    @property
    def title(self):
        return self.__title
    
    @title.setter
    def title(self, var):
        self.__title = ""
        if var:
            self.__title = self.cleantext(var)       
    
    @property
    def firstname(self):
        return self.__firstname
    
    @firstname.setter
    def firstname(self, var):
        self.__firstname = ""
        if var:
            self.__firstname = self.cleantext(var)


    @property
    def lastname(self):
        return self.__lastname
    
    @lastname.setter
    def lastname(self,var):
        self.__lastname = ""
        if var:
            self.__lastname = self.cleantext(var)

    @property
    def num_share(self):
        return self.__num_share
    
    @num_share.setter
    def num_share(self, var):
        self.__num_share = 0
        if var:
            self.__num_share = int(self.cleantext(var).replace(",", ""))

    @property
    def share_amount(self):
        return self.__share_amount

    @share_amount.setter
    def share_amount(self, var):
        self.__share_amount = 0.0
        if var:
            self.__share_amount = float(self.cleantext(var).replace(",", ""))

    @property
    def cert_date(self):
        return self.__cert_date
    
    @cert_date.setter
    def cert_date(self, var):
        self.__cert_date = ""
        if var:
            self.__cert_date = datetime.strptime(var, '%d/%m/%Y')

    @property
    def line_name(self):
        return self.__line_name

    @line_name.setter
    def line_name(self, var):
        self.__line_name = ""
        if var:
            self.__line_name = var
    
    @property
    def fb_name(self):
        return self.__fb_name

    @fb_name.setter
    def fb_name(self, var):
        self.__fb_name = ""
        if (var):
            self.__fb_name = var

    '''
    create shareholder certificate jpeg image
        @param {data} Shareholder data
        @param {string} filename to save jpeg image
        return true for success create jpeg image file
    '''
    def create_cert(self, filename):
        if not (self.no_cert and self.shareholder_id):
            return False

        if (self.num_share == 0):
            return False

        font = ImageFont.truetype(FONT_FILE, size=FONT_SIZE)

        # initial background shareholder image
        bgimg = Image.open(BGIMG_FILE)
        draw = ImageDraw.Draw(bgimg)

        # Draw text in background image
        draw.text(CERT_POS["shareholder_id"], self.shareholder_id, FONT_COLOR, font=font)
        draw.text(CERT_POS["no_cert"], self.no_cert, FONT_COLOR, font=font)
        draw.text(CERT_POS["name"], "{} {} {} ".format(self.title, self.firstname, self.lastname), FONT_COLOR, font=font)
        draw.text(CERT_POS["num_share"], "{:,}".format(self.num_share), FONT_COLOR, font=font)
        draw.text(CERT_POS["share_amount"], "{:,.2f}".format(self.share_amount), FONT_COLOR, font=font)
        draw.text(CERT_POS["share_bahttext"], "( {} )".format(bahttext(self.share_amount)), FONT_COLOR, font=font)
        draw.text(CERT_POS["cert_date"], thai_strftime(self.cert_date, "%d %B %Y"), FONT_COLOR, font=font)

        refer = ""
        if (self.line_name):
            refer = "LINE : {}".format(self.line_name)

        if (self.fb_name):
            refer =  "{}  FB: {}".format(refer, self.fb_name)
        
        #print(refer)
        draw.text(CERT_POS["refer"], refer, "rgb(255, 255, 255)", font=font)

        bgimg.save(filename)

        ## clear image
        del draw
        bgimg.close()

        return True
  
    '''
    clean space, \n, \r, \t, \u200b
    '''
    def cleantext(self, str):
        return str.strip().replace("\n", "").replace("\r", "").replace("\t", "").replace("\u200b", "")

