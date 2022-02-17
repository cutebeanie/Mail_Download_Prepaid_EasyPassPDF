'''
//// Automatic Download PDF - Receipt Easy Pass ////
Concept : E-mail provided link that need id encryted to open to print as pdf to store
Step 1 - Get Link from E-mail (by using impalib - get mail content)
Step 2 - Use Selenium to fill in ID and open the link accesing the receipt 
Step 3 - Use Selenium to print as PDF - as not option to download as pdf
'''

#Related Library for mails
import imaplib
import base64 #for reading the content of the mail
import email #get e-mail message
import re, json #Use RegX to extract link out of text in content

#Related Libraray for web access
import time
from webbrowser import get
from selenium import webdriver
import pyautogui


# //// Set login information and actuall acess in the email ////

class Login_gmail:
    login_path = r'C:\Users\admin\mail.json'

    def login_info(self):
        import json #Input for login the mail
        up = open(self.login_path)
        updata = json.load(up)
        email_user  = updata['mail']
        email_pw = updata['password']
        up.close()
        return email_user, email_pw

    def loging_in(self):
        user, pw = self.login_info()
        #get access into email - change email setting before accesing 
        # - Allow Less Secure App in google account setting 
        # - Allow IMAP access in the mail setting 
            ## Caution - less security but more convince

        mail = imaplib.IMAP4_SSL('imap.gmail.com',993)
        mail.login(user,pw)

        #if not sure can reset the setting at the end of project 
        #should do once a month for efficientcy 
        #-- water should so on 12 but pay using apps 
        #-- electricity should be done on 7 - get invoice first and then get receipt separately
        return mail
    
# //// Get the mail that we want to donwload by filtering (using search) and use the later items ////

        pyautogui.alert(text=f'Download Total of --> {numlink} links', title='Notification', button='OK') #Pop up to notify completion

# Use Selemium to fill and save pdf form link extracted
class EPprepaid_pdf:
    #Call function to work and keep as driver to be called to used later
    RT_driver = r'C:\Users\admin\chromedriver.exe' #<--- Input chrome driver location
    ch_driver = RT_driver

    mg_id = '010553001465700003'

    input_fill = '//input[@class="ant-input ng-untouched ng-pristine ng-invalid ng-star-inserted"]'
    button = '//button[@class="login-form-button login-form-margin ant-btn ant-btn-primary"]'

    def __init__(self, link_list):
        self.link = link_list

    def get_driver(self):
        #Print pdf - download locatio in Downloads
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
        prefs = {'printing.print_preview_sticky_settings.appState': json.dumps(settings)}
        chrome_options.add_experimental_option('prefs', prefs)
        chrome_options.add_argument('--kiosk-printing')
        CHROMEDRIVER_PATH = self.ch_driver
        driver = webdriver.Chrome(chrome_options=chrome_options, executable_path=CHROMEDRIVER_PATH)

        return driver

    def actiontoget_pdf(self, link):
        #Need to reopen driver everytime due to saving to pdf naming - if not reopen it will overwrite what has been saved
        driver = self.get_driver()
        
        driver.get(link)                
        time.sleep(2)
                                
        driver.find_element_by_xpath(self.input_fill).send_keys(self.mg_id)  
        time.sleep(1)

        driver.find_element_by_xpath(self.button).click()
        time.sleep(6) #May need to adjust becasue some take longer to load

        driver.execute_script('window.print();')
        
        time.sleep(2)
        driver.close()

    def get_pdf(self):
    #Action of geting all pdf download form mail
        count = 0
        for link in self.link:
            self.actiontoget_pdf(link)
            print(f'Document # {count+1} has been saved !!!')
            count +=1
            
        pyautogui.alert(text=f'Download Invoice--> {count} files completed', title='Notification', button='OK')

if __name__ == "__main__":
    pass

#///////////////////////// Action ///////////////////////////
#Get in mail
get_mailin = Login_gmail()
mail = get_mailin.loging_in()
print('Complete Get mail')

#Defind number of mails
m_num = pyautogui.prompt(text='จำนวนเมล์ที่ต้องการ Download - กรุณากรอกเป็นตัวเลข', title='Numbers#' , default='1')
num = int(m_num)


def get_mailmsg(each_mea_data):
    #Reference Step - NO USE - 
    msg = email.message_from_bytes(each_mea_data)
    content = msg.get_payload() #Get content of the e-mail

    #Decode - reference base64 library
    detail = base64.b64decode(content)
    info = detail.decode('utf-8').split()[5]

    #get link - reference regx - re Library
    link = re.findall(r'href="([\w\d:./?=&]+)"',info)[0]

    easypass['link'].append(link)

#Accessing Mail box - mail arrived in Inbox
mail.select('Inbox') #search where to search before actual search

easypass = { 'mail' : 'no-reply@etax.exat.co.th',
            'link' : []}

#Search mail using sender e-mail
mail_from = easypass['mail']
search_info = '(FROM "' + mail_from + '")'

result, data = mail.search(None, search_info)

ids = data[0] # data is a list.
id_list = ids.split() # ids is a space separated string

#Identify number of mails want to download 
set_mea = id_list[-num:] # get the latest

#Get e-mail Body - in bytes form
mea_data = []
for msg in set_mea:
    result, data = mail.fetch(msg, "(RFC822)") # fetch the email body (RFC822) for the given ID
    raw_email = data[0][1]
    mea_data.append(raw_email)
    # here's the body, which is raw text of the whole email
    # including headers and alternate payloads

#Action Getting Link - using of user defined function
for item in mea_data:
    get_mailmsg(item)
    
numlink = len(easypass['link'])

print(f'Got total of {numlink} links')
easypass['link']
print('Complete get link')

#selenium to download link - print as pdf
get_file = EPprepaid_pdf(easypass['link'])
get_file.get_pdf()
print('Complete get pdf files')
