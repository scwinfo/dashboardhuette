#---------------------------------------------------------
# Import Stuff
#---------------------------------------------------------

# Dashboarding
import streamlit as st

# Datenverarbeitung
import pandas as pd
from pandas import DataFrame
import numpy as np

# Plotting des Datenframes
import plotly.express as px
import plotly.graph_objects as go
from plotly.colors import n_colors, hex_to_rgb,label_rgb
import plotly.figure_factory as ff 

# Zeit- und Datumsberechnung
import datetime
import calendar

# Alles zum Verbinden von privatem Google-Sheet
import gspread
from google.oauth2.service_account import Credentials
from gspread_pandas import Spread, Client

# Mailing
import sib_api_v3_sdk
from sib_api_v3_sdk.rest import ApiException
from pprint import pprint
from string import Template
import time
import webbrowser

# HTML Tempalte für Mailversand
html_template = open("anfrage_huette.html").read();

configuration = sib_api_v3_sdk.Configuration()
configuration.api_key['api-key'] = st.secrets["mail_key"]

api_instance = sib_api_v3_sdk.TransactionalEmailsApi(sib_api_v3_sdk.ApiClient(configuration))


#---------------------------------------------------------
# Hüttenpreise 
#---------------------------------------------------------

ppN_e_nm = 16;
ppN_j_nm = 10;
ppN_k_nm = 6;

ppN_e_mg = 10;
ppN_j_mg = 6;
ppN_k_mg = 3;

#---------------------------------------------------------
# Verbindung zum Googlesheet herstellen
#---------------------------------------------------------

# Define the scope
scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']

# Add credentials to the account --> Local run
# credentials = Credentials.from_service_account_file('.streamlit/huettenbelegung.json', scopes=scope)

import json
key_dict = json.loads(st.secrets["text_key"])

#credentials = Credentials.from_service_account_info(key_dict)
credentials = Credentials.from_service_account_info(key_dict, scopes=scope)

# Read the data
client = Client(scope=scope, creds=credentials)

# Open Spreadsheet
sh = client.open("Hüttenbelegung")

# Open Worksheet
worksheet = sh.worksheet("Buchungen")

# Worksheet to Dataframe
df = DataFrame(worksheet.get_all_values())[1:]

# Set first row of Worksheet to column names
df.columns = df.iloc[0]
df = df[1:]

# Replace free cells to NaN
df = df.replace(r'^\s*$', np.nan, regex=True)

# Convert first column to datetime64[ns]
df["Datum"] = pd.to_datetime(df["Datum"], format="%d.%m.%Y") 

#---------------------------------------------------------
# Layout Dashboard + INPUT
#---------------------------------------------------------

# Seitentitel und Icon
st.set_page_config(page_title = "SCW-Skihütte",page_icon=":house_with_garden:")

st.image("https://www.s-c-w.com/wp-content/uploads/2022/08/logo-neu.svg", width=80)

# Subheader
st.subheader('Buchungstool Skiclub Wiebelsksirchen e.V.')

# Mainheader
st.header('Zimmerbuchung Skihütte Wiebelskirchen')

# Default Reisezeitraum (Aktueller Tag, Aktueller Tag+1)
today    = datetime.date.today()

tomorrow = today + datetime.timedelta(days=1)



# Anzeit in Dashboard
start_date = st.date_input('Anreise', today)

end_date   = st.date_input('Abreise', tomorrow)

# Fehlermeldung: Anreisedatum vor Abreisedatum
if start_date > end_date:
    st.error('Error: Abreisedatum muss hinter Anreisedatum liegen!')

# st.write(df) # ->Debugging


df_zeitraum = df[df["Datum"].isin(pd.date_range(start_date,end_date))]

#st.dataframe(df_zeitraum) # ->Debugging

if "DISPFORM" not in st.session_state:
    st.session_state.DISPFORM = 0

if "DISPKAL" not in st.session_state:
    st.session_state.DISPKAL = 0
    
if "ZIMMER" not in st.session_state:
    st.session_state.ZIMMER = []


with st.container():
    # Filterung des Zeitraum-Datenframes nach Zimmer  
    with st.form('form'):
        sel_column = st.multiselect('Zimmer auswählen:', df_zeitraum.columns[1:], 
        help='In Zeile klicken um weitere Zimmer zur Auswahl hinzufügen. Drück "Zimmer frei?" um Auwahl.')
     
        submitted = st.form_submit_button("Zimmer frei?")
        
    if submitted:
        dfnew = df_zeitraum[sel_column]
        
        # Kein Zimmer ausgewählt
        if sel_column == []:
            st.error('Kein Zimmer ausgewählt. Bitte wähle mindestens ein Zimmer.')
            st.session_state.DISPFORM = 0
        
        # Mindestens ein Zimmer wurde ausgewählt    
        else:
            
            dfnew = dfnew.fillna('frei')
            idx = (dfnew != 'frei')
            dfnew[idx] = 'Besetzt'
            
            #df_zeitraum["Datum"]
            df_zeitraum["Datum"] = df_zeitraum["Datum"].dt.strftime("%d %B %Y")
            
            #st.write(a)
            result = pd.concat([df_zeitraum["Datum"],dfnew], axis=1)
            
            sum_p = np.sum(np.array(idx))
            #st.write(sum_p) # ->Debugging
            

            if sum_p == 0:
            
                st.success('Zimmer noch frei, jetzt anfragen!')
                st.write('Belegungsplan für den gewählten Reisezeitraum')
                fig1 =  ff.create_table(result)
                st.plotly_chart(fig1)
                
                st.session_state.DISPFORM = 1
                st.session_state.DISPKAL  = 1
                                 
            else:
                st.error('Mindestens ein gewähltes Zimmer im Reisezeitraum besetzt!')
                st.write('Belegungsplan für den gewählten Reisezeitraum')
                fig1 =  ff.create_table(result)
                st.plotly_chart(fig1)
                
                st.session_state.DISPFORM = 0
                st.session_state.DISPKAL  = 0

    #https://stackoverflow.com/questions/68265188/trying-to-send-emails-from-python-using-smtplib-and-email-mime-multipart-gettin



if st.session_state.DISPFORM == 1:
    
    with st.form("my_form",):
     
        st.write("Kontaktdaten")
        vorname = st.text_input("Vorname")
        nachname = st.text_input("Nachname")
        mail = st.text_input("E-Mail")
        telefon = st.text_input("Telefon")
        naechte = (end_date - start_date).days
        
        st.write("Personenanzahl")
        c1, c2, c3 = st.columns(3)
        with c1:
            e_nm = st.number_input("Erwachsene Nichtmitglied",min_value=0, max_value=50, value=1, step=1)
            p_e_nm = e_nm*naechte*ppN_e_nm
        with c2:
            j_nm = st.number_input("Jugendliche Nichtmitglied",min_value=0, max_value=50, value=1, step=1)
            p_j_nm = j_nm*naechte*ppN_j_nm
        with c3:
            k_nm = st.number_input("Kinder Nichtmitglied",min_value=0, max_value=50, value=1, step=1)
            p_k_nm = k_nm*naechte*ppN_k_nm
        
        per_nm = str(k_nm+j_nm+e_nm)
        
        
        c4, c5, c6 = st.columns(3)
        with c4:
            e_mg = st.number_input("Erwachsene Mitglied",min_value=0, max_value=50, value=1, step=1)
            p_e_mg = e_mg*naechte*ppN_e_mg
        with c5:
            j_mg = st.number_input("Jugendliche Mitglied",min_value=0, max_value=50, value=1, step=1)
            p_j_mg = j_mg*naechte*ppN_j_mg
        with c6:
            k_mg = st.number_input("Kinder Mitglied",min_value=0, max_value=50, value=1, step=1)
            p_k_mg = k_mg*naechte*ppN_k_mg
            
        per_mg = str(k_mg+j_mg+e_mg)
        
        p_ges = str(p_e_nm + p_j_nm + p_k_nm + p_e_mg + p_j_mg + p_k_mg)
        
        LIST = [vorname,nachname,mail,telefon]
        ireadyforsubmission = "" in LIST
       
        submitted_ = st.form_submit_button("Zimmer anfragen.")
        
        zimmer = ', '.join(sel_column)

        html_template = open("anfrage_huette.html").read();
        html_mail = Template(html_template).safe_substitute(vorname=vorname,
                                                    nachname=nachname,
                                                    mail=mail,
                                                    zimmer=zimmer,
                                                    start=start_date,
                                                    ende=end_date,
                                                    per_mg=per_mg,
                                                    per_nm=per_nm,
                                                    k_mg=k_mg,
                                                    k_nm=k_nm,
                                                    j_mg=j_mg,
                                                    j_nm=j_nm,
                                                    e_mg=e_mg,
                                                    e_nm=e_nm,
                                                    p_k_mg=p_k_mg ,
                                                    p_k_nm=p_k_nm,
                                                    p_j_mg=p_j_mg,
                                                    p_j_nm=p_j_nm,
                                                    p_e_mg=p_e_mg,
                                                    p_e_nm=p_e_nm,
                                                    p_ges=p_ges
                                                   )
        html_mail_send = html_mail.encode("latin-1").decode("utf-8")
        
        if submitted_:
            
            if ireadyforsubmission == False:
                
                st.success('Kontaktdaten wurden erfolgreich übermittelt. Wir leiten Sie in ein paar Sekunden zurück auf unsere Website.')
                
                subject = f"Anfrage Skihütte Wiebeleksirchen vom {start_date} - {end_date}"
                
                #html_content = f"<h1>Vielen Dank für Ihre Anfrage</h1>"
                
                sender = {"name":"Skiclub Wiebelskirchen e.V.","email":"info@s-c-w.com"}
                space = " "
                to = [{"email":f"{mail}","name":f"{vorname+space+nachname}"}]
                cc = [{"email":"dagmar.schaufert@s-c-w.com","name":"Dagmar Schaufert"}]
                reply_to = {"email":"dagmar.schaufert@s-c-w.com","name":"Dagmar Schaufert"}
                send_smtp_email = sib_api_v3_sdk.SendSmtpEmail(to=to, html_content=html_mail_send, sender=sender, subject=subject, reply_to=reply_to)
     
                try:
                    api_response = api_instance.send_transac_email(send_smtp_email)
                    pprint(api_response)
                    
                except ApiException as e:
                    print("Exception when calling SMTPApi->send_transac_email: %s\n" % e)
             
                   
                st.session_state.DISPFORM = 0 
             
                time.sleep(2.0)
                webbrowser.open('http://s-c-w.com',new=0)
             
            else:
                st.error('Kontaktformular unvollständig. Bitte Eingabe überprüfen.')
       
 
