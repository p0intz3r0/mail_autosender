import easygui as gui
import os
import win32com.client as client
import re
import sys

def main():
    folder = gui.diropenbox() + '\\'
    if folder == None :
        main()
    mail2 = gui.enterbox("Renseignez l\'adresse e-mail")
    if check(mail2) ==0 :
        gui.msgbox("Adresse email invalide")
        main()
    else :
        msg = "Le programme enverra chaque fichier .pdf dans un mail distinct à l\'adresse selectionnee : "
        title = "Souhaitez-vous continuer ?"
        if gui.ccbox(msg + mail2, title):
            for filename in os.listdir(folder):
                if filename.endswith(".pdf"):
                    mail(filename, folder, mail2)
        else:
            sys.exit(0)


def mail(filename, folder, mail2):
    o = client.Dispatch("Outlook.Application")
    Msg = o.createItem(0)
    Msg.to = mail2
    fact_name = filename.split('.',1)[0]
    print(fact_name)
    Msg.Subject = fact_name
    Msg.Body = fact_name
    path = folder + filename
    Msg.Attachments.Add(path)
    Msg.Send()


def check(mail2):
    match = re.match('^[_a-z0-9-]+(\.[_a-z0-9-]+)*@[a-z0-9-]+(\.[a-z0-9-]+)*(\.[a-z]{2,4})$', mail2)
    if match == None:
        return 0



if __name__ == "__main__":
    main()