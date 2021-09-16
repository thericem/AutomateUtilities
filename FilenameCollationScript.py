# -*- coding: utf-8 -*-
"""
Created on Tue Sep  7 14:12:55 2021

Purpose: To walk a file tree and generate a list of files, file paths, and file extensions within the tree

A subfunction also pulls out any attachments from outlook file messages saved in the file tree and saves them separately

@author: tmorris
"""

# Function to walk a folder tree and generate a list of files and folder names, also has a function

import glob
import os
import pandas as pd
import win32com.client

# define main folder and save_name
main_folder = r"C:\ " #CHANGE AS NEEDED
save_name = 'test.csv' #CHANGE AS NEEDED


def collate_file_path_list(main_folder, save_name, extension=None):
    """
    Create a list of all files, or all files with a certain extension, within the main_folder.
    Saves the list of filenames, folders, extesions, and full-paths to a csv in the main folder.
    :param main_folder: String, path to root folder
    :param save_name: String, name for output CSV
    :param extension: String, e.g. 'pdf', 'csv'
    :return:
    """
    # Generate list of all files within any subdirectories of the main folder.
    if extension is None:
        path_list = glob.glob(main_folder+'/**', recursive=True)
    # If needed to only look for .csv files uncomment the next line and comment the previous one.
    else:
        path_list = glob.glob(main_folder+'/**/*.'+extension, recursive=True)

    # Generate list of files and folders
    file_list = []
    folder_list = []
    file_extension = []
    for i in range(len(path_list)):
        file_list.append(os.path.basename(path_list[i]))
        folder_list.append(os.path.dirname(path_list[i]))
        file_extension.append(os.path.splitext(os.path.basename(path_list[i]))[1][1:])

    # Save list to csv
    df = pd.DataFrame(data={'filename':file_list, 'folder':folder_list, 'file_extension':file_extension, 'full_path':path_list})
    df.to_csv(main_folder+'/'+save_name, index=False)
    print('Folder walk complete!')
    return


def extract_email_attachments(main_folder):
    """
    Extract and save all attachements on .msg files within any folder or subfolder of the main_folder
    :param main_folder: String, root folder of walk
    :return:
    """
    extension = 'msg'
    path_list = glob.glob(main_folder+'/**/*.'+extension, recursive=True)

    for file in path_list:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        msg = outlook.OpenSharedItem(file)  # filename including path
        att = msg.Attachments
        for i in att:
            i.SaveAsFile(os.path.join(os.path.dirname(file), (msg.Subject+'_'+i.FileName)))  # Saves the file with the attachment name

    print("Attachments extracted successfully!")
    return


extract_email_attachments(main_folder=main_folder)
collate_file_path_list(main_folder=main_folder, save_name=save_name)



