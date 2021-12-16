# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.

import os
import sys
import argparse
import logging
import jinja2

import pypff
import unicodecsv as csv
from collections import Counter


def print_hi():
    # Use a breakpoint in the code line below to debug your script.
    opst = pypff.open("C:\\Users\\Marc\\Downloads\\info@ticket-travelcenter-fuessen.com.ost")
    root = opst.get_root_folder()
    # Press Ctrl+F8 to toggle the breakpoint.
    folderTraverse(root)


def folderReport(message_list):
    """
    The folderReport function generates a report per PST folder
    :param message_list: A list of messages discovered during scans
    :folder_name: The name of an Outlook folder within a PST
    :return: None
    """
    print(message_list)

    # CSV Report


def checkForMessages(folder):
    """
    The checkForMessages function reads folder messages if present and passes them to the report function
    :param folder: pypff.Folder object
    :return: None
    """
    logging.debug("Processing Folder: " + folder.name)
    message_list = []
    for message in folder.sub_messages:
        message_dict = processMessage(message)
        message_list.append(message_dict)
    folderReport(message_list)


def processMessage(message):
    """
    The processMessage function processes multi-field messages to simplify collection of information
    :param message: pypff.Message object
    :return: A dictionary with message fields (values) and their data (keys)
    """
    return {
        "subject": message.subject,
        "sender": message.sender_name,
        "header": message.transport_headers,
        "body": message.plain_text_body,
        "creation_time": message.creation_time,
        "submit_time": message.client_submit_time,
        "delivery_time": message.delivery_time,
        "attachment_count": message.number_of_attachments,
    }


def folderTraverse(base):
    """
    The folderTraverse function walks through the base of the folder and scans for sub-folders and messages
    :param base: Base folder to scan for new items within the folder.
    :return: None
    """
    for folder in base.sub_folders:
        if folder.number_of_sub_folders:
            folderTraverse(folder) # Call new folder to traverse:
        checkForMessages(folder)


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    print_hi()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
