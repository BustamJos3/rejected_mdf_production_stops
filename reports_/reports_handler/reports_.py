# Extracted Functions
def find_reports_in_onedrive():
    """
    Scans the subfolders under the current user's OneDrive folder (including variations like 'OneDrive - Company Name')
    and returns the paths of all folders with the prefix 'reports'.

    Returns:
        list: A list of full paths to folders starting with 'reports', or an empty list if none are found.
    """
    user_home = os.path.expanduser('~')
    onedrive_folder = None
    for folder in os.listdir(user_home):
        if folder.startswith('OneDrive -'):
            onedrive_folder = os.path.join(user_home, folder)
            break
    if not onedrive_folder:
        raise FileNotFoundError('OneDrive folder not found for the current user.')
    report_folders = []
    for root, dirs, files in os.walk(onedrive_folder):
        for dir_name in dirs:
            if dir_name.lower().startswith('reports'):
                report_folders.append(os.path.join(root, dir_name))
    return report_folders

def convert_excels_to_parquet(folder_path):
    """
    Converts all Excel files in a folder to Parquet format, removes the original Excel files,
    and skips the most recent Excel file.

    Args:
        folder_path (str): Path to the folder containing Excel files.
    """
    if not os.path.exists(folder_path):
        raise FileNotFoundError(f'The folder {folder_path} does not exist.')
    excel_files = [f for f in os.listdir(folder_path) if f.endswith(('.xlsx', '.xls'))]
    if not excel_files:
        print('No Excel files found in the folder.')
        return
    full_paths = [os.path.join(folder_path, f) for f in excel_files]
    most_recent_file = max(full_paths, key=os.path.getmtime)
    most_recent_file_name = os.path.basename(most_recent_file)
    print(f'Skipping the most recent file: {most_recent_file_name}')
    for excel_file in excel_files:
        if excel_file == most_recent_file_name:
            continue
        try:
            excel_path = os.path.join(folder_path, excel_file)
            df = pd.read_excel(excel_path)
            df = df.astype(str)
            parquet_file = os.path.splitext(excel_file)[0] + '.parquet'
            parquet_path = os.path.join(folder_path, parquet_file)
            df.to_parquet(parquet_path, index=False)
            print(f'Converted: {excel_file} -> {parquet_file}')
            os.remove(excel_path)
            print(f'Removed: {excel_file}')
        except Exception as e:
            print(f'Error processing {excel_file}: {e}')

def remove_specific_chars(string=None):
    new_cause_basic_name = string
    nan_causes = 'causa no especifica'
    ignored_articles = ['de', 's', '-']
    ignored_digits = [str(number) for number in range(9 + 1)]
    for ignored_article in ignored_articles:
        if string == nan_causes:
            continue
        new_cause_basic_name = new_cause_basic_name.replace(ignored_article, '')
        new_cause_basic_name = ' '.join(new_cause_basic_name.split())
        new_cause_basic_name = ''.join([i for i in new_cause_basic_name if not i.isdigit()])
        print(new_cause_basic_name)
    return new_cause_basic_name

# Extracted Functions
def get_local_folder():
    return os.path.dirname(os.path.realpath(__file__))

# Extracted Functions
def find_files_by_format(root_folder, file_format):
    """
    Searches for all files with a specific format in a folder and its subfolders.

    Parameters:
    root_folder (str): The root folder to start the search.
    file_format (str): The file extension to search for (e.g., '.pptx', '.txt').

    Returns:
    list: A list of absolute paths to all matching files.
    """
    matching_files = []
    for root, dirs, files in os.walk(root_folder):
        for file in files:
            if file.lower().endswith(file_format.lower()):
                matching_files.append(os.path.join(root, file))
    return matching_files

def operate_mercyful_times(hour, operand):
    from datetime import datetime, timedelta
    delta_time_mercy = timedelta(minutes=0)
    mercy_flag = False
    if mercy_flag:
        delta_time_mercy = timedelta(minutes=10)
    if operand == '+' and mercy_flag:
        result_hour = eval("datetime.strptime(hour, '%H:%M:%S'){}delta_time_mercy".format(operand))
        result_hour = result_hour.time()
        return result_hour
    elif operand == '-' and mercy_flag:
        result_hour = eval("datetime.strptime(hour, '%H:%M:%S'){}delta_time_mercy".format(operand))
        result_hour = result_hour.time()
        return result_hour
    else:
        return eval("datetime.strptime(hour, '%H:%M:%S').time()")

def get_monday_of_week():
    """
    This function, `get_monday_of_week`, returns the date  
    of the Monday in the current week based on today's date.  
    
    It first retrieves today's date using `datetime.today()`  
    and then determines the current weekday using  
    `today.weekday()`, where Monday is represented by 0.  
    
    Next, it calculates the number of days to subtract to  
    reach Monday by taking the value of `today.weekday()`  
    and subtracting that number of days using `timedelta`.  
    
    Finally, it formats the result as 'YYYY-MM-DD' and  
    returns it as the output.  
    """
    today = datetime.today()
    delta = today.weekday()
    monday = today - timedelta(days=delta)
    return monday.date()

def date_limits(moment_of_week=True, day_separation=1):
    """change dates limits to filter completions 
    base on bool (default True), whether is weekend 
    or in between the week and if the moment in between 
    the week is 1 day apart or more"""
    if not moment_of_week:
        day_separation = 3
        get_current_recent_date = get_monday_of_week()
    else:
        get_current_recent_date = datetime.now()
        print(get_current_recent_date)
    day_to_stop_search = get_current_recent_date - timedelta(days=1)
    print(day_to_stop_search)
    day_before_yesterday = day_to_stop_search - timedelta(days=day_separation)
    print(day_before_yesterday)
    return (day_to_stop_search, day_before_yesterday)

def shift_checker(date_time=None, system_to_test=None):
    system = system_to_test.lower()
    print(date_time)
    time_of_interest = date_time.time()
    initial_shift_hours = shifts_hours.loc[:, [shifts_hours_cols[0]]].values
    final_shift_hours = shifts_hours.loc[:, [shifts_hours_cols[1]]].values
    available_systems = [sys[0] for sys in shifts_hours.loc[:, [shifts_hours_cols[-1]]].values]
    for available_system in available_systems:
        if system in available_system:
            print(available_system)
            matching_shift = available_systems.index(available_system) + 1
            print(matching_shift)
            if matching_shift < 4:
                print(system, ' in ', matching_shift)
                if time_of_interest >= initial_shift_hours[0][0] or time_of_interest < final_shift_hours[0][0]:
                    shift_result = 'T1'
                if time_of_interest >= initial_shift_hours[1][0] and time_of_interest < final_shift_hours[1][0]:
                    shift_result = 'T2'
                if time_of_interest >= initial_shift_hours[2][0] and time_of_interest < final_shift_hours[2][0]:
                    shift_result = 'T3'
            elif time_of_interest >= initial_shift_hours[3][0] and time_of_interest < final_shift_hours[3][0]:
                shift_result = 'T4'
            return shift_result
        else:
            print('sistema no reconocido en {}'.format(available_system))

def iter_cells(table):
    """
    to change font of text in table
    """
    for row in table.rows:
        for cell in row.cells:
            yield cell

# Extracted Functions
def aligment_paragraph(p, agliment):
    """
    the function sets the aligment for a paragraph object 
    through the string aligment, which can only be the 
    next values:
    -center
    -right
    -left
    -justify
    """
    str_valid_aligments = 'center, right, left, justify'
    try:
        a = agliment.upper()
        p_format = p.paragraph_format
        p_format.alignment = eval('WD_ALIGN_PARAGRAPH.' + a)
    except:
        if a.lower() not in str_valid_aligments:
            print(f'is {a} corresponding with')
        else:
            str_other_exceptions = 'check paragraph object valid state'
            print(str_other_exceptions)

# Extracted Functions
def image_to_base64(image_path):
    with open(image_path, 'rb') as image_file:
        base64_data = base64.b64encode(image_file.read()).decode()
    return base64_data

# Extracted Functions
def get_png_images_with_structure(root_folder):
    """
    Finds all .png images within a folder structure and creates a DataFrame
    with the root folder, successive subfolders, and the image name.

    Args:
        root_folder (str): Path to the root folder.

    Returns:
        DataFrame: A pandas DataFrame containing folder structure and image names.
    """
    data = []
    for root, dirs, files in os.walk(root_folder):
        for file in files:
            if file.endswith('.png'):
                full_path = os.path.join(root, file)
                relative_path = os.path.relpath(full_path, root_folder)
                path_parts = relative_path.split(os.sep)
                row = [root_folder] + path_parts
                data.append(row)
    max_cols = max((len(row) for row in data))
    columns = ['Root'] + [f'Subfolder_{i + 1}' for i in range(max_cols - 2)] + ['Image Name']
    df = pd.DataFrame(data, columns=columns)
    for col in df.columns[:-1]:
        df[col] = df[col].astype(str) + '\\'
    return df