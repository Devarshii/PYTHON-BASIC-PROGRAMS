import win32com.client as win32
olApp = win32.Dispatch('Outlook.Application')
olNS = olApp.GetNameSpace('MAPI')
mail_item = olApp.CreateItem(0)
mail_item.subject = "chaos of the cosmos"
mail_item.BodyFormat = 1
mail_item.body = 'chaos is primarily etiological — it relates to an attempt at explaining how the world came to exist — while “cosmos” is used as a way to describe the world as it is. Chaos was in the past, cosmos is in the present.'
mail_item.Sender = 'dtrivedi@desototechnologies.com'
mail_item.To = 'dtrivedi@desototechnologies.com'
mail_item.Save()
mail_item.Send()

