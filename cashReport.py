import win32com.client
import os
import pandas as pd
import numpy as np
from datetime import datetime
from MyFunctions import dfinfo, ScriptVars
from forex_python.converter import CurrencyRates
import glob 


def saveEmails(emailFolder,SaveFolder):
    
    mapi = outlook.GetNamespace("MAPI")
    inbox = mapi.GetDefaultFolder(6).Folders.Item(emailFolder)
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)
    message = messages.GetFirst()
    subject = message.Subject

    #Setting folder directory to save attachment
    archiveDir = os.path.join(dirname, SaveFolder)

    #Main savings script
    counter = 0

    for Item in messages:
        if counter >= 5:
            break

        for attachment in Item.attachments:
            if attachment.FileName.endswith(".png") or attachment.FileName.endswith(".msg"):
                pass
            else:
                attachment.SaveAsFile(os.path.join(archiveDir,attachment.FileName))
                print("Saving: "+ str(attachment.Filename))

        counter += 1

def readRecon(emailFolder,dropCols,acctList,dateStart,dateEnd):
    #Check to see if todays email has been sent yet, else pause script for 5 mins
    mapi = outlook.GetNamespace("MAPI")
    inbox = mapi.GetDefaultFolder(6).Folders.Item(emailFolder)
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)


    for x in messages:
        try:
            message = x
            subject = x.Subject
            df = pd.read_html(x.htmlbody)
            break 
        except ValueError:
            continue
    
    #Converting outlook email to dataframe
    df = df[0]

    #Converting date strings to date dtype before making it a column ID
    df.loc[0,dateStart:dateEnd] = df.loc[0,dateStart:dateEnd].str.replace("_","-")
    df.loc[0,dateStart:dateEnd] = pd.to_datetime(df.loc[0,dateStart:dateEnd]).dt.date

    #Renaming columns to first row, droping rows with all NA, dropping columns up to selected amounts
    #Then resetting index
    df= df.rename(columns=df.iloc[0], inplace = False)
    df= df.drop(df.index[0], inplace=False)
    df = df.dropna(how='all')
    df = df[df.columns[:dropCols]]
    df= df.reset_index(drop=True)
    
    #Can be removed once Acct ID changes
    df["Account ID"] = df["Account ID"].str.replace("_USD", "")

    #Removing accts not in account list
    df = df[df['Account ID'].isin(acctList)]

    #Reindex rows to Account and Currency
    df = df.set_index(['Account ID', 'Trade Currency'])
    df = df.sort_index()

    return df

def pullUserInput(df,fileName,startDateCol):
    tdf = pd.read_csv(os.path.join(userIntputDir,fileName),header = None)

    #Converting date strings to date dtype before making it a column ID
    tdf.iloc[0,startDateCol:] = pd.to_datetime(tdf.iloc[0,startDateCol:]).dt.date

    #Renaming columns to first row, droping rows with all NA, dropping columns up to selected amounts
    #Then resetting index
    tdf= tdf.rename(columns=tdf.iloc[0], inplace = False)
    tdf= tdf.drop(tdf.index[0], inplace=False)
    tdf= tdf.reset_index(drop=True)

    #Reindex rows to Account and Currency
    tdf = tdf.set_index(['Account ID', 'Trade Currency'])
    tdf = tdf.sort_index()

    #Converting whole dataframe to numeric and combined duplicate dates
    tdf= tdf.apply(pd.to_numeric,errors='coerce',downcast='integer')
    tdf = tdf.groupby(tdf.columns, axis=1).sum()

    #Deleting date columns that are already on the Cash Report (no double counting)
    coldel =[]
    for column in tdf.columns:
        if column in df.columns or column <= min(df.columns) :
            coldel.append(tdf.columns.get_loc(column))
    tdf= tdf.drop(tdf.columns[coldel], axis=1, inplace=False)


    #Summing all additional columns to create the additional trades column 
    tdf['Additional Trades'] = tdf.sum(axis=1, numeric_only= True)
            
    #Combining main df and the additional trades column
    df = pd.concat([df, tdf['Additional Trades']], axis=1)

    #Combining user inputted trades and recon delayed Trades 
    df= df.apply(pd.to_numeric,errors='coerce',downcast='integer')
    df['Additional Trades'] = df['Additional Trades']


    

    return df

def pullLiquidity(df,directory):
    #Finding most recent CSV file
    file_type = '\*csv'
    files = glob.glob(directory + file_type)
    max_file = max(files[-10:], key=os.path.getctime)
    print (max_file)

    keepcols = ['ID','Portfolio Name','DESCRIPTION_ADJ_LEUNGANN','LONGNAME_DESCRIPCOUPONMATDATE_LEUNGANN',
            'CS9_ADJ_LEUNGANN','CS10_ADJ_LEUNGANN','Market Value', 'MV_CORPONLY_LEUNGANN', 'Pos (Disp)',
            'Cusip  9 digits', 'Crncy']

    colsnum=list(range(0,30))
    adf = pd.read_csv(str(max_file),names=colsnum, header=None)

    #Temp variables
    idxdel1 = []
    coldel = []
    idxdel2 = []

    #Deleting blank rows on top to get column headers
    for i in adf.index:
        if (adf.iloc[i,2] is None or adf.iloc[i,2] == "NaN" 
            or pd.isnull(adf.iloc[i,2]) is True):
                idxdel1.append(i)
        else:
            break
    adf = adf.drop(adf.index[idxdel1], inplace=False)
    adf= adf.rename(columns=adf.iloc[0], inplace = False)
    adf= adf.drop(adf.index[0], inplace=False)
    adf= adf.rename(columns={adf.columns[1]:"ID"}, inplace = False)


    #Deleting columns not needed
    adf.columns = adf.columns.fillna('to_drop')
    adf= adf.drop('to_drop', axis = 1, inplace = False)
    for column in adf.columns:
        if column not in keepcols:
            coldel.append(adf.columns.get_loc(column))
    adf= adf.drop(adf.columns[coldel], axis=1, inplace=False)

    #Deleting all positions not in liquidty list
    adf = adf[adf['DESCRIPTION_ADJ_LEUNGANN'].isin(liquidityList)]

    #Removing POSONLY in portfolio names and  portfolios not in account list
    adf["Portfolio Name"] = adf["Portfolio Name"].str.replace("POSONLY", "")
    adf = adf[adf['Portfolio Name'].isin(acctList)]

    #Resetting the index
    adf= adf.reset_index(drop=True, inplace=False)

    #Setting dtypes for columns being used
    adf['Market Value']= adf['Market Value'].apply(pd.to_numeric,errors='coerce')
    adf['Crncy'] = adf['Crncy'].astype(str)
    

    #Getting USDCAD rate and applying to applicable positions and setting market value to int
    #c = CurrencyRates()
    #USDCAD = c.get_rate('USD', 'CAD')
    USDCAD = 1.351

    for j in adf.index:
        if adf.loc[j,"Crncy"] == 'USD':
            adf.loc[j,"Market Value"] = (adf.loc[j,"Market Value"]*1000)/USDCAD
        else:
            adf.loc[j,"Market Value"] = (adf.loc[j,"Market Value"]*1000)
    adf['Market Value'] = adf['Market Value'].astype(int)

    #Renaming columns
    adf = adf.rename(columns={"Portfolio Name": "Account ID"}, inplace=False)
    adf = adf.rename(columns={"Crncy": "Trade Currency"}, inplace=False)
    adf = adf.rename(columns={"LONGNAME_DESCRIPCOUPONMATDATE_LEUNGANN": "Security"}, inplace=False)
    
    #Creating pivot table
    adf = pd.pivot_table(adf, values='Market Value', index=['Account ID', 'Trade Currency'],
                         columns='Security')
    adf = adf.sort_index()

    adf['Total Liquidity'] = adf.sum(axis=1)
    adf= adf.apply(pd.to_numeric,errors='coerce',downcast='integer')

    df = pd.concat([df, adf['Total Liquidity']], axis=1)

    return df, adf

def pullLoans(df,folder,loanTotalCol):
    directory = os.path.join(dirname,folder)
    list_of_files = glob.glob(directory +'/*') # * means all if need specific format then *.csv
    max_file = max(list_of_files, key=os.path.getctime)
    print (max_file)

    ldf = pd.read_excel(str(max_file), header=None)

    #Only including rows with a acct in the name
    ldf = ldf[ldf[ldf.columns[0]].str.contains('|'.join(acctList),na=False)]
    ldf = ldf[ldf[ldf.columns[0]].str.contains('Subtotal',na=False)]

    #Resetting the ldf index
    ldf= ldf.reset_index(drop=True, inplace=False)

    #Creating Loan column in df
    df['Loans'] = np.nan

    #Applying loan total to correct account in the main df 
    for i in df.index:
        if i[1] == "USD":
            try:
                df.loc[i,'Loans'] = ldf.loc[ldf.index[ldf[ldf.columns[0]].str.contains(f'{i[0]}') == True].tolist()[0],loanTotalCol]
            except IndexError:
                continue

    #Ensuring negatives from sheet are correct, converting column to float, and flipping sign
    df['Loans'] = df['Loans'].astype(str).str.replace('\((.*)\)', '-\\1')
    df['Loans'] = df['Loans'].str.replace(",", "")
    df['Loans']= df['Loans'].apply(pd.to_numeric,errors='coerce',downcast='integer')
    df['Loans'] = df['Loans']*-1

    return df
    
def pullProjections(df,folder,projTotalCol):
    directory = os.path.join(dirname,folder)
    list_of_files = glob.glob(directory +'/*') # * means all if need specific format then *.csv
    max_file = max(list_of_files, key=os.path.getctime)
    print (max_file)

    cdf = pd.read_excel(str(max_file))

    #Only including rows with a acct in the name
    cdf['Fund #'] = cdf['Fund #'].astype(str)
    cdf = cdf[cdf['Fund #'].str.contains('|'.join(acctList),na=False)]

    #Removing rows with todays date
    cdf['Settlement'] = pd.to_datetime(cdf['Settlement'], format='%Y%m%d')
    cdf = cdf[cdf['Settlement'] != today]

    #Making fund #s and index and summing up the projections
    cdf = cdf.groupby(['Fund #'])[['Amount']].sum()

    #Creating Projection column in df
    df['Projections'] = np.nan
    for i in df.index:
        if i[1] == "CAD":
            try:
                df.loc[i,'Projections'] = cdf.loc[i[0],'Amount']
            except KeyError:
                continue 

    return df

def createACRTable(folder):
    directory = os.path.join(dirname,folder)

    #Deleting files with Calgary in the name
    for f in glob.glob(directory +'/*'):
        if 'Calgary Trading ACR' in f or 'Fund of Funds ACR' in f :
            try:
                os.remove(f)
            except OSError:
                continue
    
    list_of_files = glob.glob(directory +'/*') # * means all if need specific format then *.csv
    max_file = max(list_of_files, key=os.path.getctime)
    

    ACRdf = pd.read_excel(str(max_file), header=None)

    #Only including rows with a acct in the name
    ACRdf[1] = ACRdf[1].astype(str)
    ACRdf = ACRdf[ACRdf[ACRdf.columns[1]].str.contains('|'.join(acctList),na=False)]

    #Indexing by column 1 and sorting
    ACRdf = ACRdf.rename(columns={ACRdf.columns[1]: "Account ID"}, inplace=False)
    ACRdf = ACRdf.rename(columns={ACRdf.columns[-1]: "ACR Flow"}, inplace=False)
    ACRdf = ACRdf.set_index(ACRdf["Account ID"], inplace= False)
    ACRdf = ACRdf.sort_index()

    ACRdf = ACRdf[["ACR Flow"]]

    return ACRdf

def createETFTable(folder):
    directory = os.path.join(dirname,folder)

    #Deleting files with Calgary in the name
    for f in glob.glob(directory +'/*'):
        if 'ETF_BISSETT_CORE' in f or 'ETF_Bissett_FI_Cash_Projection' in f :
            print (f'Deleting: {f}')
            try:
                os.remove(f)
            except OSError:
                continue

    list_of_files = glob.glob(directory +'/*') # * means all if need specific format then *.csv
    max_file = max(list_of_files, key=os.path.getctime)
    print (max_file)

    etfDF = pd.read_excel(str(max_file),'FMUS', header=None)

    #Renaming columns to first row then dropping columns 
    etfDF= etfDF.rename(columns=etfDF.iloc[0], inplace = False)
    etfDF= etfDF.drop(etfDF.index[0], inplace=False)

    #Removing unneeded rows and columns
    etfDF = etfDF.iloc[0:1,2:6]

    etfDF = etfDF.set_index(etfDF["Fund/Activity"], drop=False, inplace= False)
    etfDF = etfDF.sort_index()
    etfDF = etfDF.rename(index={'FMUS':'FHIS'})

    etfDF = etfDF.drop("Fund/Activity", axis=1, inplace=False)

    etfDF = etfDF.astype(int)
    
    #print (etfDF)

    return etfDF
    

def createHTML(df,htmlName):

    df = df.copy(deep=True)

    #Converting whole df to int (no decimals) and filling Na with 0
    df = df.fillna(0)
    df = df.astype(int,errors='ignore')

    #Adjust df styles
    s = df.style

    def style_negative(v, props=''):
        return props if v < 0 else None

    border_style ='1px solid #000000 !important'
    #Adding row borders
 
    for i, _ in df.iterrows():
        if i[1] == "CAD":
            df = s.set_table_styles({i:[{'selector': 'th,td','props': [('border-top', '1px solid #000000')]}]},overwrite=False, axis=1)
            
    df = s.set_table_styles([{'selector': 'th,td','props': [('border-left', '1px solid #000000')]}],overwrite=False, axis=0)
   
        
    df = s.applymap(style_negative, props='color:red;')
    df = s.format(thousands=",")
    df = s.set_table_attributes('style="border-spacing: 1px; font: 16px;"')

    print (df)

    #write html to file
    html = df.to_html()
    
    text_file = open(os.path.join(dirname,htmlName), "w")
    text_file.write(html)
    text_file.close()



def sendEmail():
    '''
    To Do:
        Include Date checks
    '''
    file_type = '\*html'
    files = glob.glob(dirname + file_type)

    #Closing outlook
    #outlook.Quit()
    
    mail = outlook.CreateItem(0)

    #'darius.taheri@franklintempleton.ca'
    #'bissettfixedincome@franklintempleton.com'
    mail.To = 'BFIT@franklintempleton.com'
    mail.Subject = f'DRAFT Cash Flow Report - {today}'
    
    data = ""
    for i in files:
        with open(i, 'r') as myfile:
            data += myfile.read()
            data += "<br><br>"

    mail.HTMLBody = data

    dataWrite = open("W:\\CLG\\Bissett Investment Management (BIM)\\2. BIM Fixed Income Group\\..Darius\\cashReport.html","w")
    dataWrite.write(data)
    dataWrite.close()

    #mail.Display()
    mail.Send()

    #Reopening outlook
    #os.startfile("outlook")
 
if __name__ == '__main__':
    #--- Script Variables --- 
    ScriptVars()
    dirname = os.path.dirname(os.path.abspath(__file__))
    today = datetime.today().strftime('%Y-%m-%d')
    acctList = ['13758','25431','2545','2521','4431','4875']
    outlook = win32com.client.Dispatch("Outlook.Application")
    userIntputDir = r'W:\\CLG\\Bissett Investment Management (BIM)\\2. BIM Fixed Income Group\\..Darius'

    #Recon function variables
    startDateCol = 2

    #AA function variables
    AAdirectory = r'W:\\CLG\\Bissett Investment Management (BIM)\\2. BIM Fixed Income Group\\..Darius\\AAArchive'
    liquidityList =['US TREASURY N/B', 'TREASURY BILL']

    #Loan function variables 
    loanTotalCol = 8

    #Projection function variables
    projTotalCol = 4

    #Closing outlook
    print ("--- Closing Outlook ---")

 
    #--- Saving Appropriate emails from outlook ---- 
    print ("\n--- Saving Loan Emails ---")
    saveEmails("Loans", "loansArchive")

    print ("\n--- Savings Cash Projection Emails ---")
    saveEmails("Cash Projections", "projArchive")

    #--- Pull Recon and build main DataFrame ---
    print ("\n--- Pulling recon email and building main df ---")
    df = readRecon("Cash Recon",5,acctList,startDateCol,4)

    #---- Adding FHIS row ----
    print ("\n--- Creating FHIS Table ---")
    saveEmails("ETFProjections", "ETFProjArchive")
    etfDF = createETFTable("ETFProjArchive")

    #--- Pulling additional trades not reflected ---
    print ("\n--- Pulling Additional Trades Not Reflected ---")
    df = pullUserInput(df, 'CashReportUserInput.csv',startDateCol)
    
    #--- Adding liquidity data from AA sheet ---
    print ("\n--- Creating Liquidity column from AA Sheet ---")
    df,adf = pullLiquidity(df,AAdirectory)

    #---- Adding in loans column ---
    print ("\n--- Creating Loans Column from loan sheet ---")
    df = pullLoans(df,"loansArchive",loanTotalCol)

    #---- Adding in Projection column ---
    print ("\n--- Creating Projections Column from projection sheet ---")
    df = pullProjections(df,"projArchive",projTotalCol)

    

    print ("\n--- Creating Total Column ---")
    #Creating total column
    df['T+2 Balance'] = df.iloc[:,2:].sum(axis=1)

    #---- Creating ACR Table  ---
    print ("\n--- Creating ACR Table ---")
    saveEmails("ACR", "ACRArchive")
    ACRdf = createACRTable("ACRArchive")
    
    
    #Final DF formatting and HTML creation
    print ("\n--- Formatting and creating HTML files ---")
    createHTML(df,"cashreport.html")
    createHTML(adf,"liquidity.html")
    createHTML(etfDF,"ETFreport.html")
    createHTML(ACRdf,"zFlowACR.html")

    print ("\n--- Sending Email ---")
    sendEmail()
    
    print ("--- Re opening outlook ---")
    
    


    #--- Final formatting for all dataFrames
        #-- Remove decimals, add commmas, fill all zeros, NaNs with blanks
    

    #print(df)
    #print(adf)
    

    

    #Add in 

    #--- Email report with Tables ----
        #Include dates for each report used (each function return date of the report used#
        #Add in exact liquidity position df 

        #Add in ACR Section

        #Inlcude way to save the recon email for Sevrika


    
    
    '''

    - Pull recon email
    - Pull loans email
    - Pull Cash projection email
    - Pull liqudity data from AA archive
    - Pull ATNR section from user inputs
    - Check if all report dates match todays date



    '''
