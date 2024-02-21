import win32com.client as win32
def search_outlook_inbox(subject):
    outlook_app = win32.Dispatch('Outlook.Application')
    outlook_namespace = outlook_app.GetNamespace('MAPI')
    inbox_folder = outlook_namespace.GetDefaultFolder(6)
    inbox_items = inbox_folder.Items
    for item in inbox_items:
        if item.Class == 43 and item.Subject == subject:
            print(item)
            print("Received Time:", item.ReceivedTime)
            print("Sender:", item.SenderName)
            print("---------------------------------")
search_outlook_inbox("chaos of the cosmos")
