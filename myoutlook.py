# -*- coding: utf-8 -*-
"""
Created on Thu Dec 28 09:07:24 2023

@author: chanboonkwee
"""
import win32com.client

def search_email_in_folder(folder, subject: str='', method: str = 'contains', direction: str = 'latest'):
    direction = direction.strip().lower()
    direction = 'latest' if direction not in ['latest', 'first'] else direction
    method = method.strip().lower()
    method = 'contains' if method not in ['contains', 'startswith', 'endswith'] else method
    payload = {}

    items = folder.Items
    if direction == 'latest':
        items.Sort("[ReceivedTime]", True)
    for item in items:
        if any([method == 'startswith' and item.Subject.startswith(subject),
                method == 'endswith'   and item.Subject.endswith(subject),
                method == 'contains'   and subject in item.Subject]):
            print(f'"{item.Subject}" ({item.ReceivedTime}) from {item.SenderName}')
            payload[item.ReceivedTime] = item.Body
    return payload


def pull_em_text_all(subject:str='')-> list:
    if subject=='':
        return None
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # 6 refers to the index of the inbox folder

    print(f'Searching for "{subject}" in Outlook mails..')
    payload = search_email_in_folder(inbox, subject)
    for folder in inbox.Folders:
        payload.update(search_email_in_folder(folder, subject))

    return list(payload.values())


if __name__=='__main__':
    pull_em_text_all()
