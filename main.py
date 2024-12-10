# sending birthday message to friends

# import the libraries
import datetime
import pandas as pd
import smtplib 
from dotenv import load_dotenv
import os 
os.chdir(r"D:\hacker\birthday")

# Load environment variables from .env file
load_dotenv()

GMAIL_ID = os.getenv("GMAIL_ID")
GMAIL_PSWD = os.getenv("GMAIL_PSWD")


# Enter the authentication details
GMAIL_ID = "banjade7p@gmail.com"
GMAIL_PSWD = "rcdx scnj oikl lpml" #user your password in this

# function to send email

def sendEmail(to, sub, msg):
    print(f"Email to {to} sent with subjects: {sub} and message {msg}")
    s = smtplib.SMTP("smtp.gmail.com", 587, timeout=30)
    s.starttls()
    s.login(GMAIL_ID, GMAIL_PSWD)
    s.sendmail(GMAIL_ID, to, f"Subject: {sub}\n\n{msg}")
    s.quit()

if __name__ == "__main__":

    # to-test
    # sendEmail(GMAIL_ID, "subject", "test message")
    # exit()

    # read the data frame
    df = pd.read_excel("data.xlsx")
    df["Year"] = df["Year"].astype(str)
    today = datetime.datetime.now().strftime("%d-%m")

    # to write in the year, if birthday was wished
    writeInd= []
    yearNow= datetime.datetime.now().strftime("%Y")

    # iterate through the dataframe
    for index, item in df.iterrows():
        bday = item["Birthday"].strftime("%d-%m")

        # comapare the birth-date with today's date
        if (today == bday) and yearNow not in str(item["Year"]):
            sendEmail(item["Email"], "Happy Birthday", item["Dialogue"])
            writeInd.append(index)

    for i in writeInd:
        yr = df.loc[i, "Year"]
        df.loc[i, "Year"] = str(yr) + "," + str(yearNow)

        df.to_excel("data.xlsx", index=False)