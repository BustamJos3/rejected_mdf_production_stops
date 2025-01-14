#!/usr/bin/env python
# coding: utf-8

# In[3]:


import os
from pathlib import Path
import pandas as pd


# In[4]:


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


# In[5]:


# export df as excel data
folder_path=r"C:\Users\JDBUSTAMANTE\OneDrive - Duratex SA\reports_visualizacion_data_produccion\source_return_data\data_plots\imgs_reports_daily"
df_images = get_png_images_with_structure(folder_path)
df_images.to_excel(r"C:\Users\JDBUSTAMANTE\OneDrive - Duratex SA\reports_visualizacion_data_produccion\source_return_data\path_structure_of_plots.xlsx")


# In[6]:


get_ipython().system('jupyter nbconvert --to script plot_paths_to_pwbi_report.ipynb')

