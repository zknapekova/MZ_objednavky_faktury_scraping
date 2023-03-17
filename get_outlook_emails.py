import win32com.client  # pywin32 package needs to be installed
import os
import pandas as pd


class OutlookTools:
    def __init__(self, object):
        self.obj = object
        self.n_folders = self.obj.Folders.Count

    def show_all_folders(self):
        for i in range(0, self.n_folders):
            print(f'Folder: [{i}] {self.obj.Folders[i].Name}')
            n_subfolders = self.obj.Folders[i].Folders.Count
            for j in range(n_subfolders):
                print(f'    Subfolder: [{j}] {self.obj.Folders[i].Folders[j].Name}')
                if self.obj.Folders[i].Folders[j].Folders.Count != 0:
                    for k in range(self.obj.Folders[i].Folders[j].Folders.Count):
                        print(f'        Subfolder: [{k}] {self.obj.Folders[i].Folders[j].Folders[k].Name}')

    def find_message(self, folder_path: str, condition: str):
        '''

        :param folder_path: example - outlook.Folders['zuzana.knapekova@health.gov.sk'].Folders['Doručená pošta']
        :param condition: possible filters to use: subject, sender, to, body, receivedtime etc.
        :return: item object with filtered messages
        '''
        messages_all = folder_path.Items
        return messages_all.Restrict(condition)

    def save_attachment(self, output_path, messages):
        '''

        :param output_path: folder for saving attachments
        :param messages: item object containing at least one message
        :return:
        '''
        for message in messages:
            ts = pd.Timestamp(message.senton).strftime('%d_%m_%Y_%I_%M_%S')
            for attachment in message.Attachments:
                try:
                    print(message.senton)
                    attachment.SaveASFile(os.path.join(output_path, ts+'_'+attachment.FileName))
                    print(f"attachment {attachment.FileName} from {message.Sender} saved")
                except Exception as e:
                    print("error when saving the attachment:" + str(e))
        print('All atachments were saved')

# Example of use:
#path = outlook.Folders['zuzana.knapekova@health.gov.sk'].Folders['Doručená pošta']
#OT = OutlookTools(outlook)
#result = OT.find_message(path, "[SenderName] = 'Gajdošová Denisa'")
#OutlookTools(outlook).save_attachement(os.getcwd(), result)

# Useful links:
# another way of getting inbox folder - https://learn.microsoft.com/en-us/office/vba/api/outlook.oldefaultfolders
# all available properties for message - https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook.mailitem
# restriction guide - https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook._items.restrict
# advanced filtering - https://learn.microsoft.com/en-us/office/vba/outlook/how-to/search-and-filter/filtering-items-using-query-keywords

