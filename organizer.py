#!/usr/bin/env python
# coding: utf-8

# In[12]:


import os
from pathlib import Path
import json
from jupyter_core.paths import jupyter_runtime_dir
from base_fns import get_local_folder
import ast


# In[13]:


class repo_organizer:
    def __init__(self):
        self.notebook_path=self._get_notebook_path()
        self.current_workdir=self._set_working_directory()
        self.python_scripts=self._file_format_searcher(format_str="*.py")
        
    def _get_notebook_path(self):
        """
        Get the absolute path of the currently running Jupyter Notebook by finding the active kernel.
    
        Returns:
            str: The absolute path of the notebook or an error message.
        """
        try:
            rt_fldr = get_local_folder()
            return rt_fldr
    
        except Exception as e:
            return str(e)
            
    def _set_working_directory(self):
        """
        Set the working directory to the specified path.

        Args:
            path (str): The path to set as the working directory.

        Returns:
            str: The updated working directory path or an error message.
        """
        try:
            os.chdir(self.notebook_path)
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
            os.system(f"cmd < {cmd_file}")

        except Exception as e:
            print(f"An error occurred: {e}")

    def _file_format_searcher(self,format_str="*.none"):
        """
        Search for all files in the current working dir and generate an iterable with their paths
        """
        cwd = Path(os.getcwd())
        return cwd.rglob(format_str)

    def _execute_matching_scripts(self):
        """
        Execute all Python scripts in the subfolders of the current working directory if their names match
        the name of a Jupyter notebook (.ipynb) within the same folder, using the Python instance
        of the current virtual environment on Windows.

        Returns:
            None
        """
        try:
            # Get the current working directory
            cwd = Path(os.getcwd())

            # Find all Jupyter notebooks and Python scripts in the subfolders
            notebooks = {notebook.stem for notebook in self._file_format_searcher("*.ipynb")}
            scripts = {script for script in self.python_scripts}

            # Match Python scripts with Jupyter notebook names
            matching_scripts = [script for script in scripts if script.stem in notebooks]

            if not matching_scripts:
                print("No matching scripts found.")
                return

            # Get the Python executable of the current environment (Windows)
            venv_dir = os.environ.get("VIRTUAL_ENV")
            if venv_dir:
                python_executable = Path(venv_dir) / "Scripts" / "python.exe"
            else:
                python_executable = Path(os.sys.executable)  # Fallback to the default Python executable

            if not python_executable.exists():
                raise FileNotFoundError(f"Python executable not found: {python_executable}")

            # Execute each matching script
            for script in matching_scripts:
                print(f"Executing {script}...")
                os.system(f"{python_executable} {script}")

        except Exception as e:
            print(f"An error occurred: {e}")
            
    def _export_requirements(self):
        """
        Generate a requirements.txt file for the Python scripts in the current repo folder,
        explicitly searching for .py files and using pipreqs to analyze dependencies.

        Returns:
            None
        """
        try:
            import os
            from pathlib import Path

            # Determine the current working directory (repo folder)
            repo_dir = Path(os.getcwd())

            # Ensure the virtual environment's Scripts directory contains pipreqs
            venv_dir = os.environ.get("VIRTUAL_ENV")
            if venv_dir:
                pipreqs_path = Path(venv_dir) / "Scripts" / "pipreqs.exe"
            else:
                raise FileNotFoundError("Virtual environment not found. Activate a venv before running this.")

            if not pipreqs_path.exists():
                raise FileNotFoundError(f"pipreqs not found at {pipreqs_path}")

            # Run pipreqs explicitly on the repo folder
            requirements_in = repo_dir / "requirements.in"
            os.system(f"{pipreqs_path} --force --savepath={requirements_in}")

            # Compile requirements.in into requirements.txt
            pip_compile_path = Path(venv_dir) / "Scripts" / "pip-compile.exe"

            if not pip_compile_path.exists():
                raise FileNotFoundError(f"pip-compile not found at {pip_compile_path}")

            os.system(f"{pip_compile_path} {requirements_in}")

            print("Requirements export completed.")

        except Exception as e:
            print(f"An error occurred: {e}")

    def _extract_functions_to_module(self, input_file, output_module):
        """
        Extract all function definitions from a Python file and save them to a new module,
        avoiding duplicate functions already present in the module.

        Args:
            input_file (str): Path to the input Python file.
            output_module (str): Path to the output module file.

        Returns:
            None
        """
        try:
            import ast
            from pathlib import Path

            # Read the input file
            with open(input_file, "r") as f:
                source_code = f.read()

            # Parse the source code into an AST
            tree = ast.parse(source_code)

            # Extract function definitions
            functions = [
                node for node in tree.body if isinstance(node, ast.FunctionDef)
            ]

            if not functions:
                print("No functions found in the input file.")
                return

            # Check existing functions in the output module
            existing_functions = set()
            if Path(output_module).exists():
                with open(output_module, "r") as f:
                    existing_code = f.read()
                existing_tree = ast.parse(existing_code)
                existing_functions = {
                    node.name for node in existing_tree.body if isinstance(node, ast.FunctionDef)
                }

            # Generate source code for new functions
            new_functions = [
                func for func in functions if func.name not in existing_functions
            ]

            if not new_functions:
                print("No new functions to add.")
                return

            extracted_code = "\n\n".join(
                ast.unparse(func) for func in new_functions
            )

            # Append the new functions to the output module
            with open(output_module, "a") as f:
                f.write("\n\n# Extracted Functions\n")
                f.write(extracted_code)

            print(f"Extracted functions saved to {output_module}")

        except Exception as e:
            print(f"An error occurred: {e}")


# In[14]:


organizer=repo_organizer()


# In[16]:


for python_script in organizer.python_scripts:
    organizer._extract_functions_to_module(python_script,r".\reports\reports.py")

