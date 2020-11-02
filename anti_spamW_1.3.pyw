import win32com.client
import tkinter as tk
from tkinter import ttk
from datetime import datetime, timedelta

app_version = '1.3'

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

debug = False

switcher = {
    '12': '2007',
    '14': '2010',
    '15': '2013',
    '16': '2016/2019/365' }

version = switcher[outlook.Application.Version.split('.')[0]]

words_whitelist = ['avast', 'plaisio', 'kolossos', 'travelling', '.gr/', 'youtube']
senders_whitelist = ['mailman-bounces@lists', '@tee.gr', '@europlan.gr', '@central.tee.gr', '.gov.gr', '@yahoo', '@gmail.com', '@hotmail.com']

init_days = '3'
clean_from = ( datetime.utcnow() - timedelta( days=int(init_days)) ).strftime('%Y-%m-%d')

##inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder
##junkFolder = outlook.GetDefaultFolder(23)

def debug():
    accounts = [i for i in outlook.Session.Accounts]
    
##    for acc in accounts:
##        print ('Name: ', acc.DisplayName)
##        print ('Iol: ', acc.IOlkAccount)
##        print ('Deliv: ', acc.DeliveryStore)
##        print ('Ses: ', acc.Session)
##        print ('Type: ', acc.AccountType)
##        print ('done')
##        print (acc.GetAddressEntryFromID)
##    for acc in accounts:
##        root_folder = outlook.DeliveryStore.Folders.Item(1)
##        for folder in root_folder.Folders:
##            print (folder.Store)
##    
##    # print all folders with indexes
##    for i in range(50):
##        try:
##            box = outlook.GetDefaultFolder (i)
##            name = box.Name
##            print(i, name)
##        except:
##            pass

##    print (outlook.AccountSelector.SelectedAccount)

    print ([i for i in accounts[1].Session.Folders])
    
def clean_inbox(event): #Spams without attachments, with fake links inBody
    clean_from = ( datetime.utcnow() - timedelta( days=int(entry.get()) ) ).strftime('%Y-%m-%d')
    selected = choices.index(variable.get())

    try:
        account = outlook.Session.Accounts[selected]
        inbox = account.DeliveryStore.GetDefaultFolder(6) # "6" refers to the index of a folder
        junkFolder = account.DeliveryStore.GetDefaultFolder(23)
    except: #pywintypes.com_error (beacause of opening and closing outlook)
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        account = outlook.Session.Accounts[selected]
        inbox = account.DeliveryStore.GetDefaultFolder(6) # "6" refers to the index of a folder
        junkFolder = account.DeliveryStore.GetDefaultFolder(23)
##        text_box.insert(tk.END, "\n" + 'Error! Cannot find the specified account. Please, restart the program...\nNext time, please do not open and close the Outlook while I am being run!' )
##        text_box.see("end")
##        return
##    inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder
##    junkFolder = outlook.GetDefaultFolder(23)
        
    
    
    text_box.insert(tk.END, "\nCleaning %s inbox from %s ... please wait...\n" % (choices[selected], clean_from))
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)
    i = 0
    spams_num = 0
    messageIDs = []
    message = messages.GetFirst()
    
    while str(message.ReceivedTime) >= clean_from: #message!=None:
        i+=1
##        if i==10: break
        
        rec_time = message.ReceivedTime
        if ('http://' in message.body or 'https://' in message.body) \
           and not any([whiteword in message.SenderEmailAddress for whiteword in senders_whitelist]) \
           and message.Size < 40000:
            
            try:
                htmlBody = message.HTMLBody.split('<body')[1] #'<BODY' ????            
                tagA = htmlBody.split('a href="http')[1].split('</a>')[0]
                parts = tagA.split('//')
                part1 = parts[1].split('/')[0]
                part2 = parts[2].split('/')[0]
##                if 'mutons' in message.SenderEmailAddress: print(part1, part2) #debug
                if part1 != part2 and not any([whiteword in part1 for whiteword in words_whitelist]):
                    log_string = '%s | %s (%s)' % (str(rec_time).split('+')[0], message.SenderEmailAddress, message.Sender)
                    
                    text_box.insert(tk.END, "\n" + log_string)
                    window.update()
                    text_box.see("end")
##                    message.Move(junkFolder)
                    # store message IDs to delete them later...
                    messageIDs.append(message.EntryID)
##                    spams_num+=1
            except IndexError:
##                print(rec_time, message.SenderEmailAddress)
                pass

        message = messages.GetNext()

    if debug != True:        
        for message_id in messageIDs:
            outlook.Session.GetItemFromID(message_id).Move(junkFolder)

    text_box.insert( tk.END, "\n\nDone... %d emails processed since %s, %d emails moved to Junk!!!\n" % ( i, clean_from, len(messageIDs) ) )
    text_box.see("end")
        

if __name__ == "__main__":
##    debug()
    if debug == True:
        print ('Debug on...')
    
    window = tk.Tk()
    window.title("Outlook Clean from Spam! ver. %s" % app_version) 
    label_date = tk.Label(text = 'Search for how many days: ')
    label_date.grid(row=1, column=1, sticky='e')
    entry = tk.Entry(width=4)
    entry.insert(0, init_days)    
    entry.grid(row=1, column=2, sticky='w')

##    label_acc = tk.Label(text='Using Account: ') #%s' % outlook.Session.DefaultStore)
##    label_acc.grid(row=3,column=1)  
    #Choose account
    choices = [account.DisplayName for account in outlook.Session.Accounts]
    variable = tk.StringVar(window)
    variable.set(choices[0])
    w = tk.OptionMenu(window, variable, *choices)
    w.grid(row=1, column=0, sticky='w')

    #entry = tk.Entry(fg="yellow", bg="black", width=20)
    #entry.insert(0, outlook.Session.DefaultStore)
    button = tk.Button(
        text="Clean Inbox!",
        font=("Helvetica", 14, "bold"),
        width=10,
        height=2,
        bg="blue",
        fg="yellow",
    )
    button.bind("<Button-1>", clean_inbox)
    button.grid(row=1, column=3, rowspan=2)
    
    text_box = tk.Text()
    text_box.grid(row=4, column=0, columnspan=4, sticky='nwse')
    window.grid_rowconfigure(4,  weight =1)
    window.grid_columnconfigure(0,  weight =1)


    label_ver = tk.Label( text = 'Outlook Version: %s' % version )
    label_ver.grid(row=5, column=0, sticky='w')
    
    label_licence = tk.Label(
        text='Manolis Stratigis @ TEE/TAK',
        fg='grey'
        )
    label_licence.grid(row=5, column=3, columnspan=4, sticky='nwse')

###make grid responsive
##    n_rows = 5
##    n_columns = 4
##    for i in range(n_rows):
##        window.grid_rowconfigure(i,  weight =1)
##    for i in range(n_columns):
##        window.grid_columnconfigure(i,  weight =1)
    
    window.mainloop()
