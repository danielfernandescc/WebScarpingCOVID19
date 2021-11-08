#!/usr/bin/env python
# coding: utf-8

# In[1]:


import gspread
import pandas as pd
import requests
from oauth2client.service_account import ServiceAccountCredentials
import os 
import glob

def to_dataframe(dir,painel_indexes, vacinas_indexes):
    #Open a Writer
    writer = pd.ExcelWriter('C:\\Users\\danie\\Documents\\Taxa_Normalidade\\Webscraping\\tmp\\compilado_final.xlsx', engine='xlsxwriter')
    
    #Write painel according ndexes 
    for i in painel_indexes:  
        df = pd.read_excel(dir+'saved_spreadsheet0.xlsx',sheet_name=i)
        if i == 0:
            df.to_excel(writer,sheet_name='CONFIRMADOS')
        if i == 1:
            df.to_excel(writer,sheet_name='OBITOS')
    i = 0
    
    #Write vacina according indexes
    for i in vacinas_indexes:
        df = pd.read_excel(dir+'saved_spreadsheet1.xlsx',sheet_name=i)
        print('>>',i)
        if i == max(vacinas_indexes):
            df.to_excel(writer,sheet_name='XLSX_Vacinas')
            writer.save()
            print("[TO_DATAFRAME]: OK")
        else:
            print('>>>',i)
            df.to_excel(writer,sheet_name="{}".format(i))
            
def api_web(sheets_urls):
    for i in range(len(sheets_urls)):
      resp = requests.get(sheets_urls[i])
      sheet_name = 'saved_spreadsheet' + str(i) + '.xlsx'
      #file_path = '/tmp/' + sheet_name
      file_path = "C:\\Users\\danie\\Documents\\Taxa_Normalidade\\Webscraping\\tmp\\" + sheet_name
      os.makedirs(os.path.dirname(file_path), exist_ok=True)
      output = open(file_path, 'wb')
      output.write(resp.content)
      output.close()
      print('[API_WEB] saved:',sheet_name)
      #to_dataframe()
    print("[API WEB] Write content local DONE")

def create_keyfile_dict():
    json_dict = {
        "type": "service_account",
        "project_id": "esuspainel-326311",
        "private_key_id": "2b94e51daed190def7e2ebe013c42f3bf34a08ae",
        "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvAIBADANBgkqhkiG9w0BAQEFAASCBKYwggSiAgEAAoIBAQCIPO1xVLkT/AyO\n3r3YnYYRah5sBCAH4/kOM6UEpzxHkBLjV302Y+kiWPg++mVJfVKGLcs2eBkXhvHl\nxXP+lMDpvCYhs16e8Gq2lUXXoNdAGqyQ3vd4YOGEni3orWvvcFp9gaJzgJRk9yZ4\nLp1OjkQNw960Sev34RLBO2jYufoZDk+t9P1OKVTNlTdtNd1Du6XBSh3/LNTGEjcn\nIXrQayp/QiGph5qg6PDmhd7jLqnpjIrsI8W6Oww/LLlJMoDLRgmVAXCcgy205mY9\n8kAgJhpFn+vw9jW9e8udb1ebizN50LDpwKPhtcqFRsQj0P0NobShjBpraHCr7V0z\nf+bOj4EBAgMBAAECggEAQjCbycp4SvnTnhwgz1uk9dQBaMhOSZcceyZjP5ICquAY\nFST9/A1piJsCLRLZX+2HyRH5n4KU6kXRQ6l9dAwQd97GBeyQBZdXuVJnxt3phkcP\nXSk+wVkMaDKzqk6LWJ7VEBIJ+6TWNAGRyqUXH0HmVWu6yQv7HYjX5FK5W0Zr+Hyy\niLHx7/IpkJ47Yu/z1TaRYXlnVhuOJ7kVLwDyL9ZvrUNz/g+DMM1zpfa1yEC/9g+Y\nQWlcdIgZ9SARX/pz2IV5AkDAgtm7xyPaGda44cj2DktB8nJMSHmoUms0M1Ne0pFi\nZYgWQfPHV9oskFFlouHGS4pN95cjCAWpKiRC5VnE6wKBgQC7TjoP0fxvCSqyfCiO\nK5S7r1WZmvi9E5E4neTGFsGr6aC9f2zJ0YlDf9q7IdiuKiYtiPniv4tiseEq1dRo\nDO03VtIYmCVOO8NRMKOz078AGsPIy0p1txQ0nbNj3EELWEB/Rv88vy38DPEi2SAO\nShDlBuOuK2oH+I1cfvj8dI/gMwKBgQC6NA+JTXY086IQalqYlVQns3Qu2VUI34fI\n7Z3N7bh4D/ifv9XD9qgJ1Z5gyq+sN31v8JVlQ/eMeBr2VQJ3IE3+vCBoQdbRmd8E\nXP9Y9mdBmhiJUqMWzrEBYW7VtZKazJT/hMUoxpeJ4Ab2bZ0P3ucTVMjqsGSoXFf7\nNRWaMvuV+wKBgC0GbvqilbXzVCo3omAapdRAH6mfETASZhRgEEB18/RpYtRqrzIM\nhpyNPX1Cc53aT/ceOEODm/QLon7zi+2/Pb7RxgtXd5BI2XjI4nE183II/Qtlou6N\nJfRH/HmC1rftbQOrg2uM4Xb3fXfNDeGheFI1x8F0ejaUTxbvBtdZBcT1AoGANEoS\nYthh7ZTNWhbDwj2NGGkIo29ctdUv6Hjx67ZqKy0xAIt6mEFYBwr6IuxIUPB0RU8m\nZP2lMsk3qR1OR+3GeVaTMzPqA4pWWn9TJcRsUrvXUBjou6rngh++ZD1NIjN5VBgQ\n1daPD6Tdz64QgThzY7ZXhbBrU+w6uMy7eEYA6KkCgYB17ma/DdMZS8VBAH+nIdq2\npB8mjHlKz0S4MIJv6gY8JR/pqIdPw/rP+4ZG4rq9NnLsviuiq2WFTuCfPx5ae24s\nEgFXGhUdrxbiPffQ0Y5IRMzq3+09CCn85I0TT6ojKic+CMwxa/YsSGdr+pRnRP0G\nTKQxPFjemlNtzkZbfG3I+g==\n-----END PRIVATE KEY-----\n",
        "client_email": "esus-painel-teste@esuspainel-326311.iam.gserviceaccount.com",
        "client_id": "104319409684833211694",
        "auth_uri": "https://accounts.google.com/o/oauth2/auth",
        "token_uri": "https://oauth2.googleapis.com/token",
        "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
        "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/esus-painel-teste%40esuspainel-326311.iam.gserviceaccount.com"
    }
    return json_dict

def acess_crendentials(json):
    # define the scope]
    scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets', 
         "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
  # add credentials to the account       
    creds = ServiceAccountCredentials.from_json_keyfile_dict(json, scope) 
    client = gspread.authorize(creds) 
    return client

def main():
  sheets_urls = [
                'http://sescloud.saude.mg.gov.br/index.php/s/ZEzzC8jFpobXGjM/download?path=%2FPAINEL_COVID&files=XLSX_Painel.xlsx', #painel
                'http://sescloud.saude.mg.gov.br/index.php/s/ZEzzC8jFpobXGjM/download?path=%2FVACINAS&files=XLSX_Vacinas.xlsx' #vacinas                 
                ]
  sheets_names = ['saved_spreadsheet0','saved_spreadsheet1']
  painel_indexes = [0,1]
  vacinas_indexes = [0,1,2,3,4,5,6,7,8,9]
  json = create_keyfile_dict()
  client = acess_crendentials(json)
  #api_web(sheets_urls)
  #update_spreadsheets(client,sheets_names)
  dir = 'C:\\Users\\danie\\Documents\\Taxa_Normalidade\\Webscraping\\tmp\\'
  to_dataframe(dir,painel_indexes,vacinas_indexes)
  
if __name__ == "__main__":
    main()


# In[ ]:




