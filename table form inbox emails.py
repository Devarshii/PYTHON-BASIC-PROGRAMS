import win32com.client as win32
import pandas as pd
def search_outlook_inbox():
    outlook_app = win32.Dispatch('Outlook.Application')
    outlook_namespace = outlook_app.GetNamespace('MAPI')
    inbox_folder = outlook_namespace.GetDefaultFolder(6) #inbox #5 sent
    inbox_items = inbox_folder.Items
    mail_items = []
    for item in inbox_items:
        if item.Class == 43:
            mail_items.append({
                "Subject": item.Subject,
            })
    if mail_items:
        df = pd.DataFrame(mail_items)
        html_table = df.to_html(index=False)
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.Subject = "Total inbox email list"
        mail.HTMLBody = "<html><body>" + html_table + "</body></html>"
        mail.Recipients.Add("dtrivedi@desototechnologies.com")  # Replace with recipient's email address
        mail.Send()
        print("Mail sent successfully!")
    else:
        print("No mail items found.")
search_outlook_inbox()
