import email
import imaplib
import ctypes
import getpass
import win32com.client as wincl
import speech_recognition as sr
import nltk
import os
speak = wincl.Dispatch("SAPI.SpVoice")

speak.Speak("WELCOME TO THE GMAIL ASSISTANT ")
speak.Speak("ENTER THE DETAILS")
mail = imaplib.IMAP4_SSL('imap.gmail.com',993)
user = raw_input("Enter the Gmail address : ")
pwd = getpass.getpass()
try:
    mail.login(user,pwd)
    speak.Speak("login success")
    os.system('cls')
except:
    speak.Speak("You have some problem logging to the account , check your internet connection or check whether the IMAP is enabled in the gmail... or..check the email id and the password is correct")
    exit()
mail.select("INBOX")
def look():
    mail.select("INBOX")
    n=0
    (retcode,messages) = mail.search(None,'(UNSEEN)')
    if retcode == 'OK':
        for new in messages[0].split():
            n=n+1
            typ,data = mail.fetch(new,'(RFC822)')
            for respon_part in data:
                if isinstance (respon_part,tuple):
                    original = email.message_from_string(respon_part[1])
                    froms = original['From']
                    data = original['Subject']
                    f1 = nltk.word_tokenize(froms)
                    
                    speak.Speak("you just got an email from..." + f1[0] +".." + f1[1] + "..from the email id ..." + f1[3]+f1[4]+f1[5]+"..." + "..with the subject...." + data +"..." + "..please check out...")
                   
                    
                    
                    typ,data = mail.store(new,'+FLAGS','\\Seen')
                    


if __name__ == '__main__':
    while True:
        look()
