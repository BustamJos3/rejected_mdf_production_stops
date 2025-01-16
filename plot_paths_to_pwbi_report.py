#!/usr/bin/env python
# coding: utf-8

# In[5]:


import os
from pathlib import Path
import pandas as pd


# In[6]:


def get_png_images_with_structure(root_folder):
    """
    Finds all .png images within a folder structure and creates a DataFrame
    with the root folder, successive subfolders, and the image name.

    Args:
        root_folder (str): Path to the root folder.

    Returns:
        DataFrame: A pandas DataFrame containing folder structure and image names.
    """
    # List to store data
    data = []

    # Walk through all subdirectories and files
    for root, dirs, files in os.walk(root_folder):
        for file in files:
            if file.endswith('.png'):
                # Compute relative path components
                full_path = os.path.join(root, file)
                relative_path = os.path.relpath(full_path, root_folder)
                path_parts = relative_path.split(os.sep)

                # Append data: root_folder + path parts
                row = [root_folder] + path_parts
                data.append(row)

    # Determine max number of columns for subfolders
    max_cols = max(len(row) for row in data)

    # Create column names dynamically
    columns = ['Root'] + [f'Subfolder_{i+1}' for i in range(max_cols-2)] + ['Image Name']

    # Create DataFrame
    df = pd.DataFrame(data, columns=columns)
    for col in df.columns[:-1]:  # Excluye la última columna
        df[col] = df[col].astype(str) + '\\'
    return df


# In[7]:


import os

def find_reports_in_onedrive():
    """
    Scans the subfolders under the current user's OneDrive folder (including variations like 'OneDrive - Company Name')
    and returns the paths of all folders with the prefix 'reports'.

    Returns:
        list: A list of full paths to folders starting with 'reports', or an empty list if none are found.
    """
    # Get the base path to the user's home directory
    user_home = os.path.expanduser("~")

    # Find the OneDrive folder (handles variations like "OneDrive - Company Name")
    onedrive_folder = None
    for folder in os.listdir(user_home):
        if folder.startswith("OneDrive -"):
            onedrive_folder = os.path.join(user_home, folder)
            break

    if not onedrive_folder:
        raise FileNotFoundError("OneDrive folder not found for the current user.")

    # Search for folders with the prefix 'reports' in the OneDrive directory
    report_folders = []
    for root, dirs, files in os.walk(onedrive_folder):
        for dir_name in dirs:
            if dir_name.lower().startswith("reports"):
                report_folders.append(os.path.join(root, dir_name))

    return report_folders


# In[2]:


reports_paths=find_reports_in_onedrive()
reports_paths


# In[8]:


str_folder_searcher="reports_visualizacion_data_produccion"
for report_path in reports_paths:
    if str_folder_searcher in report_path:
        path=Path(reports_paths[reports_paths.index(report_path)])
path=Path.joinpath(path,r"source_and_return_data")
path


# In[10]:


# export df as excel data
folder_path=Path.joinpath(path,r"data_plots\imgs_reports_daily")
df_images = get_png_images_with_structure(folder_path)
path_export=Path.joinpath(path,"path_structure_of_plots.xlsx")
df_images.to_excel(path_export)


# In[11]:


get_ipython().system('jupyter nbconvert --to script plot_paths_to_pwbi_report.ipynb')

