#!/usr/bin/env python
# coding: utf-8

# In[36]:


import os
from pathlib import Path
import json
from jupyter_core.paths import jupyter_runtime_dir


# In[64]:


class repo_organizer:
    def __init__(self):
        self.notebook_path=self._get_notebook_path()
        
    def _get_notebook_path(self):
        """
        Get the absolute path of the currently running Jupyter Notebook by finding the active kernel.
        
        Returns:
            str: The absolute path of the notebook or an error message.
        """
        try:
            # Locate Jupyter's runtime directory
            runtime_dir = Path(jupyter_runtime_dir())
            kernel_files = list(runtime_dir.glob("kernel-*.json"))
        
            if not kernel_files:
                raise FileNotFoundError(f"No kernel connection files found in {runtime_dir}.")
        
            # Get the current kernel ID from the environment variable
            kernel_id = os.path.basename(os.getenv("JPY_PARENT_PID", ""))
        
            # Match the active kernel file
            kernel_file = next((kf for kf in kernel_files if kernel_id in kf.stem), None)
            if not kernel_file:
                raise FileNotFoundError(f"No matching kernel file found for ID {kernel_id}.")
        
            # Read the kernel file
            with open(kernel_file, "r") as file:
                kernel_data = json.load(file)
        
            # Retrieve the notebook path if available
            notebook_path = kernel_data.get("notebook_path", "Notebook path not found in kernel file")
            return str(Path(notebook_path).resolve())
        
        except Exception as e:
            return str(e)
            
    def _set_working_directory(self,path):
        """
        Set the working directory to the specified path.

        Args:
            path (str): The path to set as the working directory.

        Returns:
            str: The updated working directory path or an error message.
        """
        try:
            os.chdir(path)
            return os.getcwd()
        except Exception as e:
            return str(e)
    def _export_notebooks_to_scripts(self):
        """
        Search for all Jupyter notebooks (.ipynb) in the subfolders of the current working directory,
        excluding files in .ipynb_checkpoints, generate CMD commands to export them as Python scripts,
        write those commands to a file called notebook_to_script_convertion.txt, and execute the file.

        Returns:
            None
        """
        try:
            # Get the current working directory
            cwd = Path(os.getcwd())

            # Find all Jupyter notebooks in the subfolders, excluding .ipynb_checkpoints
            notebooks = [
                notebook for notebook in cwd.rglob("*.ipynb")
                if ".ipynb_checkpoints" not in notebook.parts
            ]
            notebook_paths = [notebook.relative_to(cwd) for notebook in notebooks]

            if not notebook_paths:
                print("No Jupyter notebooks found in the subfolders.")
                return

            # File to store the CMD commands
            cmd_file = cwd / "notebook_to_script_convertion.txt"

            # Write CMD commands to the file
            with open(cmd_file, "w") as f:
                for notebook in notebook_paths:
                    cmd = f"jupyter nbconvert {notebook} --to script\n"
                    f.write(cmd)

            print(f"Commands written to {cmd_file}")

            # Execute the file using CMD
            os.system(f"cmd /c {cmd_file}")

        except Exception as e:
            print(f"An error occurred: {e}")


# In[65]:


organizer=repo_organizer()


# In[66]:


organizer.notebook_path


# In[67]:


organizer._export_notebooks_to_scripts()

