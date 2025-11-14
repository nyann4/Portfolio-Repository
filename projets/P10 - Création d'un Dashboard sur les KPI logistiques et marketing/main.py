import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import email
import base64
import pandas as pd
from google.oauth2 import service_account
import gspread
import warnings
import flask
import functions_framework


def clean_df_1(df, file_date): #Synthèse
    #We can keep all the file except the "total" row and the "representant"/"code client" column
    df = df.loc[df["Code client"] !="Total"]
    df = df.loc[df["En %"]!="-"]
    df = df.filter(items=['Raison sociale client',
       'Portefeuille de commandes non livrées', 'CA RAL en retard', 'En %',
       'Dt CA RAL retard de préparation', 'Dt CA RAL en retard Rupture',
       'Dt CA RAL en retard commande complète lié rupture',
       'Dt CA RAL en retard BLOCAGE COMPTA', 'Dt CA RAL en retard BLOCAGE ADV',
       'CA Cde à expédier de J+1 à J+5', ' CA Cde à expédier de J+6 à J+16',
       'CA Cde à expédier > à J+16'])
    df["date"] = file_date
    
    return df

def clean_df_2(df1, file_date): #Synthèse rupture
    #We can keep all the file except the "total" row
    df1 = df1.loc[df1["Fournisseur"] !="Total"].copy()
    df1["date"] = file_date
    return df1

def clean_df_3(df2, file_date): #Detail rupture
    #need to only keep the article that are actually missing
    nombre_article = len(df2)
    df2 = df2.loc[(df2["Nombre de lignes impactées"] != 0) & (df2["Code article"] != "Total")].copy()
    df2["date"] = file_date
    df2.loc[:,"Date de livraison fournisseur prévue"] = df2["Date de livraison fournisseur prévue"].astype(str)
    return df2, nombre_article

def clean_df_4(df3, file_date): #Detail
    #We only need the article that are missing and the date of the breakdown, we can filter all the other columns
    list_col_df_3 = ["Code client", "Raison sociale client","Code article", "Date de commande","CA Commandes en cours", "Dt CA RAL en retard",	"Dt CA RAL en retard Rupture",
                     "Dt CA RAL en retard BLOCAGE COMPTA",  "Dt CA RAL en retard BLOCAGE ADV","Dt CA RAL en retard commande complète lié rupture",
                     "CA Commandes en cours échues <5J", "CA Commandes en cours échues entre 6 et 16J", "CA Commandes en cours échues >16J"]
    df3 = df3.loc[(df3["Etat rupture"] == "En rupture") | (df3["Etat ligne commande"] == "Attente")]
    df3 = df3.filter(items=list_col_df_3).copy()
    df3["date"] = file_date
    df3["Date de commande"] = df3["Date de commande"].astype(str)
    return df3

def connect():
    creds = None
    SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first time.
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open("token.json", "w") as token:
            token.write(creds.to_json())
    return creds

def get_service(creds):
    try:
        # Call the Gmail API
        service = build("gmail", "v1", credentials=creds)
    except HttpError as error:
        # TODO(developer) - Handle errors from gmail API.
        print(f"An error occurred: {error}")
        service = "false"
    return service

# find all the messages informations that match our subject "keyword"
def search_messages(service, keyword):
    search_id = service.users().messages().list(userId="me",q=f"subject:{keyword},has:attachment").execute()
    #keep a list of dict containing the message id
    messages_dict = search_id["messages"]
    #retrieve only the messages id
    list_id = [values["id"] for values in messages_dict ]
    print(list_id)
    return list_id

#retrieve what contain a specific message
def get_message(service, msg_id):
    # return a dictionnary with multiple information related to our message (date, test, size...)
    message_list = service.users().messages().get(userId="me", id=msg_id, format="raw").execute()
    #decode the raw of the message
    msg_raw = base64.urlsafe_b64decode(message_list["raw"].encode('UTF-8'))
    print(msg_raw[0:50])
    #we convert the mail into a message object
    mime_msg = email.message_from_bytes(msg_raw)
    #the payload will be the attachment of the file
    
    for att in mime_msg.get_payload() : # verifying the type of the mail
        print(att.get_filename())
        if att.get_content_type() == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
            if att.get_filename().find("RAL_Journalier") != -1: #verifying that the xlsx file is the good one
                attachment = att
    # try :
    file = attachment.get_payload(decode="UTF-8")
    file_date = pd.read_excel(io=file, sheet_name=['Synthèse'], nrows=1)["Synthèse"]['Origine rapport - Qlik NPrinting - RAL - Journalier'].copy().values[0]
    df_synth = pd.read_excel(sheet_name=['Synthèse'], io=file, skiprows=4)['Synthèse']
    df_synth_rup = pd.read_excel(sheet_name=['Synthèse Rupture'], io=file, skiprows=4)['Synthèse Rupture']
    df_det_rupt = pd.read_excel(sheet_name=['Détail Rupture'], io=file, skiprows=4)['Détail Rupture']
    df_det = pd.read_excel(sheet_name=['Détail'], io=file)['Détail']
    return df_synth, df_synth_rup, df_det_rupt, df_det, file_date[-10:]
    # except :
    #     print("there is no xlsx attachment in this mail")

def api_gsheet_connect(): #connect to the gsheet api
    scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

    credentials = service_account.Credentials.from_service_account_file("key.json", scopes=scopes)

    file = gspread.authorize(credentials)
    RAL_timeseries = file.open("RAL TimeSeries") #open the time series file 
    #get a sheet object for each sheet of the file
    sheet_synth = RAL_timeseries.worksheet("Synthèse")
    sheet_synth_rupt= RAL_timeseries.worksheet("Synthèse Rupture")
    sheet_det_rupt = RAL_timeseries.worksheet("Détail Rupture")
    sheet_det = RAL_timeseries.worksheet("Détail")

    sheet_logs = RAL_timeseries.worksheet("logs")
    print("its working !")

    return sheet_synth, sheet_synth_rupt, sheet_det_rupt, sheet_det, sheet_logs

def from_sheet_to_df(sheet): #get the dataframe related to the sheet object
    sheet_values = sheet.get_all_values()
    df = pd.DataFrame(sheet_values[1:], columns = sheet_values[0])
    return df

@functions_framework.http
def scrapral(request: flask.Request) -> flask.typing.ResponseReturnValue:
    warnings.simplefilter(action='ignore', category=FutureWarning)
    alpha = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z']
    sheet_synth, sheet_synth_rupt, sheet_det_rupt, sheet_det, sheet_logs = api_gsheet_connect() #api gsheet dataframe creation
    #retrieve the dates from the log file, sort them ascending and get the last one
    df_date = from_sheet_to_df(sheet_logs)
    print(df_date)
    creds = connect()
    service = get_service(creds) #get service object of the gmail api
    list_id = search_messages(service, "[Qlik] reste à livrer") #check for each mail that contain "Reste à livrer"
    list_id.reverse()
    number_of_mail = 10
    for id in list_id[-number_of_mail:]  :
    # for id in list_id  :
    # we use list_id[0] to get the first item of the list which will be the last one receive
        df_1, df_2, df_3, df_4, file_date = get_message(service, id) #write the xlsx file
        #read the 4 dataframe from each of the file sheet
        df_1 = clean_df_1(df_1, file_date)
        df_2 = clean_df_2(df_2, file_date)
        df_3, nombre_article = clean_df_3(df_3, file_date)
        df_4 = clean_df_4(df_4, file_date)

        df_synth = from_sheet_to_df(sheet_synth)
        df_synth_rupt = from_sheet_to_df(sheet_synth_rupt)
        df_det_rupt = from_sheet_to_df(sheet_det_rupt)
        df_det = from_sheet_to_df(sheet_det)
        df_date = from_sheet_to_df(sheet_logs)
        if file_date not in df_date["date"].values :
            if len(df_synth) >0 : #recheck if the value inside the file is already added to the time series
                idx = len(df_synth)
                sheet_synth.update(df_1.values.tolist(), f'A{idx+1}:{alpha[len(df_1.columns)-1]}{len(df_1)+idx}')
                idx = len(df_synth_rupt)
                sheet_synth_rupt.update(df_2.values.tolist(), f'A{idx+1}:{alpha[len(df_2.columns)-1]}{len(df_2)+idx}')
                idx = len(df_det_rupt)
                sheet_det_rupt.update(df_3.values.tolist(), f'A{idx+1}:{alpha[len(df_3.columns)-1]}{len(df_3)+idx}')
                idx = len(df_det)
                sheet_det.update(df_4.values.tolist(), f'A{idx+1}:{alpha[len(df_4.columns)-1]}{len(df_4)+idx}')
            else :
                #creation d'entête de column
                sheet_synth.update([df_1.columns.tolist()], f'A{1}:{alpha[len(df_1.columns)-1]}{len(df_1)+1}')
                sheet_synth_rupt.update([df_2.columns.tolist()], f'A{1}:{alpha[len(df_2.columns)-1]}{len(df_2)+1}')
                sheet_det_rupt.update([df_3.columns.tolist()], f'A{1}:{alpha[len(df_3.columns)-1]}{len(df_3)+1}')
                sheet_det.update([df_4.columns.tolist()], f'A{1}:{alpha[len(df_4.columns)-1]}{len(df_4)+1}')

                sheet_synth.update(df_1.values.tolist(), f'A{2}:{alpha[len(df_1.columns)-1]}{len(df_1)+1}')
                sheet_synth_rupt.update(df_2.values.tolist(), f'A{2}:{alpha[len(df_2.columns)-1]}{len(df_2)+1}')
                sheet_det_rupt.update(df_3.values.tolist(), f'A{2}:{alpha[len(df_3.columns)-1]}{len(df_3)+1}')
                sheet_det.update(df_4.values.tolist(), f'A{2}:{alpha[len(df_4.columns)-1]}{len(df_4)+1}')
            idx = len(df_date)
            sheet_logs.batch_update([{'range': f'A{idx+2}','values': [[file_date]]}])
            sheet_logs.batch_update([{'range': f'B{idx+2}','values': [[nombre_article]]}])
            print(f"Ajout du fichier du {file_date} réussi")
        else :
            print(f"Ce fichier RAL à déjà été ajouté, date du fichier : {file_date}")
    return "Its done !"