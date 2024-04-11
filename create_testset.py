# Essentials
import json
import os
from pathlib import Path
import re

# For spreadsheet
import pandas as pd
from pandas import DataFrame
from xlrd import open_workbook
import tempfile
import csv
pd.options.mode.chained_assignment = None  # default='warn'

# For environment variables
from dotenv import load_dotenv

# For selenium
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.common import exceptions as selenium_exceptions
from selenium import webdriver

import time
import selenium.webdriver.common.action_chains as AC
import sys

DEBUG = True
SYQT_HOME_DIR = ""


class IncompleteColumnError(Exception):
    """Exception raised for errors in the input excel sheet.

    Attributes:
        message -- explanation of the error
    """

    def __init__(self, message="Column in the excel sheet is incomplete"):
        self.message = message
        super().__init__(self.message)

class IncorrectLoginCredentials(Exception):
    """ Exception raised for incorrect CodeBeamer credentials. """
    def __init__(self, message="Incorrect credentials used for CodeBeamer!"):
        self.message = message
        super().__init__(self.message)

class CodeBeamerMaintenance(Exception):
    """ Exception raised for CodeBeamer maintenance. """
    def __init__(self, message="CodeBeamer is currently down at the moment."):
        self.message = message
        super().__init__(self.message)

class NoEntryFound(Exception):
    """Exception raised when no entry is found using .loc of a dataframe"""
    def __init__(self, message="No specified row is found."):
        self.message = message
        super().__init__(self.message)


# Define the template for the configuration
config_template = {
    "settings": {
        "mib3oigp": {
            "components": {
                "nav": {
                    "res_col_id": "",
                    "anchor_column": "",
                    "test_case_link": "http://vwavncb.lge.com:8080/cb/tracker/73303386?view_id=-11&subtreeRoot=26867344",
                    "test_set_link": "http://vwavncb.lge.com:8080/cb/category/77502497"
                },
                "sds": {
                    "res_col_id": "",
                    "anchor_column": "",
                    "test_case_link": "http://vwavncb.lge.com:8080/cb/tracker/73303386?view_id=-11&subtreeRoot=21733926",
                    "test_set_link": "http://vwavncb.lge.com:8080/cb/category/100046653"
                }
            }
            ,
            "version_pattern": "N\\d{3}\\.\\d{2}",
            "test_set_tracker": "Test Set_VW",
            "test_run_item_prefix": "[MIB3 GP ST] TestRun_VW_NAR_",
            "test_configuration":"MIB3GP"
        },
        "nar classic": {
            "components": {
                "nav": {
                    "res_col_id": "",
                    "anchor_column": "",
                    "test_case_link": "http://vwavncb.lge.com:8080/cb/category/94014605?view_id=-2",
                    "test_set_link": "http://vwavncb.lge.com:8080/cb/category/100046653"
                },
                "sdars": {
                    "res_col_id": "",
                    "anchor_column": "",
                    "test_case_link": "http://vwavncb.lge.com:8080/cb/category/94014605?view_id=-2",
                    "test_set_link": "http://vwavncb.lge.com:8080/cb/category/100046653"
                },
                "navi + rvc": {
                    "res_col_id": "",
                    "anchor_column": "",
                    "test_case_link": "http://vwavncb.lge.com:8080/cb/category/94014605?view_id=-2",
                    "test_set_link": "http://vwavncb.lge.com:8080/cb/category/100046653"
                },
                "sds": {
                    "res_col_id": "H",
                    "anchor_column": "Name",
                    "test_case_link": "http://vwavncb.lge.com:8080/cb/category/94014605?view_id=-2",
                    "test_set_link": "http://vwavncb.lge.com:8080/cb/category/100046653"
                }
            },
            "version_pattern": "N\\d{3}\\.\\d{2}",
            "test_set_tracker": "Test Set_MQBClassicNAR",
            "test_run_item_prefix": "[MQB ClassicNAR ST] TestRun_VW_",
            "test_configuration":"MQB_NAR"
        }
        
    }
}

def load_env():
    global SYQT_HOME_DIR
    global CB_ID
    global CB_PASS

    
    if Path(r"C:\Users\lgeuser\Documents\programs\.env").exists():
        load_dotenv(dotenv_path=r"C:\Users\lgeuser\Documents\programs\.env")
        SYQT_HOME_DIR = os.getenv("SYQT_HOME_DIR")
    else:
        load_dotenv()
        SYQT_HOME_DIR = os.getenv("SYQT_HOME_DIR")
    
    CB_ID = os.getenv("CB_ID")
    CB_PASS = os.getenv("CB_PASS")
    if not (SYQT_HOME_DIR and CB_ID and CB_PASS) :
        raise Exception("Ooof, you need to configure your environment variable 😗 Ask Seung for one he hasn't given you one.")
    
def get_all_keys(nested_dict):
    keys_list = []
    for key, value in nested_dict.items():
        keys_list.append(key)
        if isinstance(value, dict):
            keys_list.extend(get_all_keys(value))
    return keys_list

def load_config():
    try:
        with open('test_set.json', 'r') as config_file:
            t = json.load(config_file)
            if get_all_keys(t) != get_all_keys(config_template):
                print(set(get_all_keys(config_template)) - set(get_all_keys(t)))
                raise json.JSONDecodeError("Invalid CONFIG file.", "", 0 )
            return t
    except (FileNotFoundError, json.JSONDecodeError):
        print("Invalid CONFIG file. Creating a new one for you...")
        # Return a copy of the template without the values
        return json.loads(json.dumps(config_template))

def save_config(CONFIG):
    with open('test_set.json', 'w') as config_file:
        json.dump(CONFIG, config_file, indent=4)

def valid_test_link(link:str)->bool:
    """ Returns true if it's a valid test set link.  """
    return link.startswith("http://vwavncb.lge.com") or\
           link.startswith("https://vwavncb.lge.com")

def get_test_link() -> str:
    while True:
        link = input()
        if valid_test_link(link): return link
        print("Enter a valid test set link.")

def update_config(CONFIG):
    project_presets: str
    product: str
    component: str
    value: str
    for product, product_settings in CONFIG['settings'].items():
        for project_presets, component_props in product_settings.items():
            if project_presets == "components":
                for component, component_preset in component_props.items():
                    for key, value in component_preset.items():
                        if key == "test_case_link":
                            if not valid_test_link(value):
                                print(f"Enter a valid test set link for {product.upper()}: ")
                                value = get_test_link()
                                component_preset["test_case_link"] = value
                        if not value:
                            # Prompt the user for the missing value
                            value = input(f"Enter the {key.replace('_', ' ')} for {product.upper()} - {component.upper()}: ")
                            # Update the configuration with the user input
                            CONFIG['settings'][product]["components"][component][key] = value
    save_config(CONFIG)

def get_user_selection(CONFIG) -> tuple[str, str, bool]:
    """ Returns selection in the format: [project name, component name, create_test_set_ans] 
    [0] proj name
    [1] component name
    [2] option to create test sets based on user input
    """
    # if DEBUG:
    #     return ["nar classic", "sdars", True]
    #     return ["mib3oigp", "sds", True]
    #     return ["nar classic", "nav", True]
    #     return ["mib3oigp", "nav", True]
    
    # Helper function to display choices and get user's selection
    def get_choice(options, prompt):
        print(prompt)
        option: str
        for i, option in enumerate(options, 1):
            print(f"{i}. {option.upper()}")
        choice = input("Select an option by number: ")
        return choice

    def ask_to_create_test_set(prompt):
        return input(prompt).strip().lower()
        
    # Get project selection
    projects = list(CONFIG['settings'].keys())
    project_prompt = "Please select a project:"
    project_choice = get_choice(projects, project_prompt)
    
    # Validate project selection
    while not project_choice.isdigit() or \
            int(project_choice) not in range(1, len(projects) + 1):
        print("Invalid selection. Please select a valid number.")
        project_choice = get_choice(projects, project_prompt)
    project = projects[int(project_choice) - 1]

    # Get component selection
    components = list(CONFIG['settings'][project]['components'].keys())
    component_prompt = f"Please select a component for {project}:"
    component_choice = get_choice(components, component_prompt)
    
    # Validate component selection
    while not component_choice.isdigit() or int(component_choice) not in range(1, len(components) + 1):
        print("Invalid selection. Please select a valid number.")
        component_choice = get_choice(components, component_prompt)
    component = components[int(component_choice) - 1]

    # Ask if user would like to also create a test set
    create_test_set = ask_to_create_test_set(
        "Do you also want to create a test set for this project? (y/n): ")

    while not (create_test_set == 'y' or create_test_set == 'n'):
        create_test_set = ask_to_create_test_set(
            "Invalid selection. Type 'y' or 'n': "
            )
    create_test_set = True if create_test_set == 'y' else False
    return (project, component, create_test_set)

def get_excel(proj_name:str, component_name: str):
    """ Returns a dataframe """
    target_file_name = f"[{proj_name}][{component_name}] SyQT Test case full.xlsx".lower()

    target_path = ""
    # Search recursively for the file
    for filename in Path(SYQT_HOME_DIR).rglob(f"*.xlsx"):
        if filename.name.lower() == target_file_name:
            print(f"Detected file: {filename.name.title()}")
            target_path = filename
            break
    else:
        raise FileNotFoundError(
            f"I couldn't find the required SyQT file 😢.\n"\
            "The file should be located in '{SYQT_HOME_DIR}' with the "\
            "following format: "\
                "'[<model name>][<component name>] SyQT Test case full.xlsx'.")
    try:
        df = pd.read_excel(target_path, keep_default_na=False)
        
    except ValueError as e:
        with tempfile.TemporaryDirectory() as tmpdir:
            wb = open_workbook(target_path)
            ws = wb.sheet_by_index(0)
            with open(f"{tmpdir}/{target_path.name}", "w", newline='') as file:
                writer = csv.writer(file)
                for row_num in range(ws.nrows):
                    # print(ws.row_values(row_num)[16])
                    writer.writerow([d for d in ws.row_values(row_num)])
            df = pd.read_csv(f'{tmpdir}/{target_path.name}', encoding="unicode_escape")
    return df

def setup_df(df: DataFrame, proj_name:str, component_name: str) -> tuple[DataFrame, str]:
    """ Cleans the dataframe on project basis and returns the version column \
        name on the sheet. For example, `N401.02 <description>`.\
        The following projects are filtered by empty corresponding column \
        value. 
        For NAR Classic:
            * NAV | Steps.Expected result
    """
    # Check if the spreadsheet has the essential columns
    df.columns = [col.lower() for col in df.columns]
    required_columns = ["comments", "name", "id"]
    for c in required_columns:
        if c not in df.columns:
            raise IncompleteColumnError(
                f"You need to include the column '{c}' in the spreadsheet."
            )
        
    # Check if every row contains a valid result
    # If not, inform the user to finish the result and quit the program
    res_col = ""
    for col_name in df.columns:
        if df[col_name].dtype == "object":
            for column_v in df[col_name][:10]:
                try:
                    t_v = column_v.lower().strip()
                    if "pass" == t_v or "fail" == t_v \
                    or "blocked" == t_v or "na" == t_v \
                    or "hold" == t_v or "excl" == t_v:
                        print(f"Detected results column: '{col_name.title()}'")
                        res_col = col_name
                        break
                except AttributeError as e:
                    # Ignore NaNs
                    continue
            else:
                continue
            break
        else:
            IncompleteColumnError("Couldn't find the appropriate result column. Check your spreadsheet!")
       
    # Ensure that the results column name contains the version number
    pattern = re.compile(CONFIG["settings"][proj_name]["version_pattern"])
    if len(res_col) < 7:
        raise IncompleteColumnError("Your results column needs to include full version info.")
    version_info = res_col[:7].upper()
    if not pattern.match(version_info):
        raise IncompleteColumnError(
            f"The column name should start with the version number in the "\
            "following format: "\
            "{CONFIG['settings'][proj_name]['version_pattern'][0]}XXX.XX"
        )
    #TODO

    # Filter all the rows that do not have a NAME by using the
    # anchor column
    anchor_col = CONFIG["settings"][proj_name]["components"][component_name]\
                       ["anchor_column"].lower()
    if anchor_col not in df.columns:
        while True:
            t = input("Invalid anchor column. Enter a valid anchor column: ").lower()
            if t in df.columns:
                CONFIG["settings"][proj_name]["components"][component_name]\
                      ["anchor_column"] = anchor_col = t
                save_config(CONFIG)
                break
            
    df.dropna(axis=0, subset=anchor_col, inplace=True)
    df.reset_index(drop=True, inplace=True)
    
    def filter_comments(d): 
        text = str(d).strip()
        if len(text) == 0:
            return 'nan'
        else: 
            return text
        
    def check_id(d):
        id = str(d).strip().split('.')[0]
        if id.isdigit():
            return id
        else:
            raise IncompleteColumnError(
                f"The ID contains non-numerical values. Please fix this. "\
                "The id: {id}")

    # Perform additional house keeping, cleaning data, etc
    # Verify if the number of items seem correct.
    df["name"] = df["name"].apply(lambda d:str(d).lower().strip())
    df["comments"] = df["comments"].apply(lambda d:filter_comments(d))
    df[res_col] = df[res_col].apply(lambda d:str(d).lower())
    df["id"] = df["id"].apply(lambda d:check_id(d))
    complete_set = set(["pass", "fail", "blocked", "na", "excl"])

    err_msg = ""
    u = set(df[res_col].unique())
    if not u.issubset(complete_set):
        raise IncompleteColumnError(
            "The result column doesn't appear to be complete. Please check this 👀.\n"\
            "You should only include pass, fail, blocked, na, excl values in "\
            "your results column."
                                    )
    res_n = df[res_col].count()
    filter_col_n = df[anchor_col].count()
    name_n = df["name"].count()
    id_n = df["id"].count()
    
    if not (res_n == filter_col_n == name_n == id_n):
        err_msg = "The number of results, names, id, and expected results "\
                  "column do not match!"
        if name_n < filter_col_n or name_n < res_n:
            err_msg += "\nThe name column appears to be missing some entries."
        raise IncompleteColumnError(err_msg)
    
    # Filter by value of the column, res_col, where it's not excluded
    df = df.loc[df[res_col] != "excl"]

    # Check for dupliate 'name' entries
    df["name_dup"] = df["name"].duplicated("first")
    df_dup = df["name"].loc[df["name_dup"] == True]

    # If duplicates exist, warn the user.
    if df_dup.shape[0] > 0:
        print("⚠ WARNING: There are duplicate test cases in this spreadsheet. ")
        print(df_dup)
        for name in df_dup.unique():
            t_df = df.loc[df["name"] == name]
            if t_df[res_col].unique().size > 1 or \
               t_df["comments"].unique().size > 1:
                raise IncompleteColumnError(f"The duplicate test cases should have the same results and the comments. Test case: {name}")

    # Ensure all failed test cases include a comment with a link
    # Ensure all blocked and failed test cases include a comment. Failed test cases 
    df_blocked = df.loc[(df[res_col] == "blocked") & (df["comments"] == 'nan')]
    if df_blocked.shape[0] > 0:
        print("Blocked TCs are missing associated comments.")
        print(df_blocked["name"])
        raise IncompleteColumnError

    df_fails = df.loc[df[res_col] == "fail"]
    fails = []
    for row in df_fails.iterrows():
        if len(url_extractor(row[1]["comments"])) == 0:
            fails.append(row[1]["name"])
    if len(fails) > 0:
        print("You are missing ticket links for failed tc's. 😢")
        for i, x in enumerate(fails,1):
            print(f"{i}. {x}")
        raise IncompleteColumnError
     
    print(f"Total number of test cases (after filtering) 🔬: {df.shape[0]}")  
    return (df, res_col)

def configure_webdriver() -> webdriver.Chrome:
    options = webdriver.ChromeOptions()
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    return webdriver.Chrome(options=options)

def recursive_search_incl_get_attr(
        driver: webdriver.Chrome, selector: str, text:str,
        node_type: str, attr: str
        ) -> tuple[WebElement, str] | None:
    """ Self explanatory, but you must only use this when the target element has an ID attribute. """
    rec_search_script = f"""
    function recSearchTextTC(elem, text, nodeType) {{
        if ((elem?.nodeName.toLowerCase() === nodeType ) && elem?.textContent.trim().toLowerCase().includes(text.toLowerCase())
        ) {{
            return [elem, elem.getAttribute("{attr}")]
        }}
        for (let child of elem?.children){{
            let t_elem = recSearchTextTC(child, text, nodeType)
            if (t_elem !== null) {{
                return t_elem
            }}
        }}
        return null
    }}
    let elem = document.querySelector("{selector}");
    return recSearchTextTC(elem, "{text}", "{node_type}");
    """
    # TODO
    return driver.execute_script(rec_search_script)


def recursive_search_incl_get_attr_excl_class(
        driver: webdriver.Chrome, selector: str,
        text:str, node_type: str, attr: str, excl_class:str
        ) -> tuple[WebElement, str] | None:
    """ Self explanatory, but you must only use this when the target element has an ID attribute. This will skip elements with a class `excl_class` """
    
    rec_search_script = f"""
    function recSearchTextTC(elem, text, nodeType) {{
        if ((elem?.nodeName.toLowerCase() === nodeType ) 
        && elem?.textContent.trim().toLowerCase().includes(text.toLowerCase())
        && ! elem.classList.contains('{excl_class}')
        ) {{
            return [elem, elem.getAttribute("{attr}")]
        }}
        for (let child of elem?.children){{
            let t_elem = recSearchTextTC(child, text, nodeType)
            if (t_elem !== null) {{
                return t_elem
            }}
        }}
        return null
    }}
    let elem = document.querySelector("{selector}");
    return recSearchTextTC(elem, "{text}", "{node_type}");
    """
    return driver.execute_script(rec_search_script)

def str_to_int(s):
    if s is None:
        return None
    return int(s)

def recursive_search_includes(
        driver: webdriver.Chrome, selector: str,
        text: str, node_type: str
        ) -> WebElement:
    """ You must use Double quotes to wrap around your selector! """

    rec_search_script = f"""
    function recSearchText(elem, text, nodeType) {{
        if (elem.nodeName.toLowerCase() === nodeType && elem.textContent.toLowerCase().includes(text.toLowerCase())) {{
            return elem;
        }}
        for (let child of elem.children){{
            let t_elem = recSearchText(child, text, nodeType);
            if (t_elem !== null) {{
                return t_elem;
            }}
        }}
        return null;
    }}
    let elem = document.querySelector("{selector}");
    return recSearchText(elem, "{text}", "{node_type}");
    """
    return driver.execute_script(rec_search_script)


def recursive_search_includes_click_js(
        driver: webdriver.Chrome, selector: str,
        text: str, node_type: str
        ) -> WebElement:
    """ You must use Double quotes to wrap around your selector! """

    rec_search_script = f"""
    function recSearchText(elem, text, nodeType) {{
        if (elem.nodeName.toLowerCase() === nodeType && elem.textContent.toLowerCase().includes(text.toLowerCase())) {{
            return elem;
        }}
        for (let child of elem.children){{
            let t_elem = recSearchText(child, text, nodeType);
            if (t_elem !== null) {{
                return t_elem;
            }}
        }}
        return null;
    }}
    let elem = document.querySelector("{selector}");
    recSearchText(elem, "{text}", "{node_type}").click();
    """
    return driver.execute_script(rec_search_script)


def recursive_search_exact_set_attr(
        driver: webdriver.Chrome, selector: str, text: str, node_type='option', attr_key='selected', attr_val=''
        ) -> WebElement:
    """ You must use Double quotes to wrap around your selector! """

    rec_search_script = f"""
    function recSearchTextExact(elem, text, nodeType) {{
         if ((elem?.nodeName.toLowerCase() === nodeType ) && (elem?.textContent.trim().toLowerCase() === text.toLowerCase())) {{
            elem.setAttribute("{attr_key}", "{attr_val}")
            return elem;
        }}
        for (let child of elem.children){{
            let t_elem = recSearchTextExact(child, text, nodeType);
            if (t_elem !== null) {{
                return t_elem;
            }}
        }}
        return null;
    }}
    let elem = document.querySelector("{selector}");
    return recSearchTextExact(elem, "{text}", "{node_type}");
    """
    return driver.execute_script(rec_search_script)



def recursive_search_exact(
        driver: webdriver.Chrome, selector: str,
        text: str, node_type: str
        ) -> WebElement:
    """ You must use Double quotes to wrap around your selector! """

    rec_search_script = f"""
    function recSearchTextExact(elem, text, nodeType) {{
         if ((elem?.nodeName.toLowerCase() === nodeType ) && (elem?.textContent.trim().toLowerCase() === text.toLowerCase())) {{
            return elem;
        }}
        for (let child of elem.children){{
            let t_elem = recSearchTextExact(child, text, nodeType);
            if (t_elem !== null) {{
                return t_elem;
            }}
        }}
        return null;
    }}
    let elem = document.querySelector("{selector}");
    return recSearchTextExact(elem, "{text}", "{node_type}");
    """
    return driver.execute_script(rec_search_script)



def select_filter(
        driver: webdriver.Chrome, text: str, node_type: str
        ) -> WebElement:
    """ You must use Double quotes to wrap around your selector. You can use this to add as many as filters as you want. The follow up selection needs to have to be customized, however. """

    rec_search_script = f"""
    function recSearchTextExact(elem, text, nodeType) {{
         if ((elem?.nodeName.toLowerCase() === nodeType ) && (elem?.textContent.trim().toLowerCase() === text.toLowerCase())) {{
            return elem;
        }}
        for (let child of elem.children){{
            let t_elem = recSearchTextExact(child, text, nodeType);
            if (t_elem !== null) {{
                return t_elem;
            }}
        }}
        return null;
    }}
    let elem = document.querySelectorAll('.ui-widget-header.ui-corner-all.ui-multiselect-header.ui-helper-clearfix.ui-multiselect-hasfilter + ul')[3];
    recSearchTextExact(elem, "{text}", "{node_type}").click();
    """
    return driver.execute_script(rec_search_script)


def select_status(
        driver: webdriver.Chrome, text: str, 
        node_type: str, toggle_not: bool
        ) -> WebElement:
    """ You must use Double quotes to wrap around your selector! """

    rec_search_script = f"""
    function recSearchTextExact(elem, text, nodeType) {{
         if ((elem?.nodeName.toLowerCase() === nodeType ) && (elem?.textContent.trim().toLowerCase() === text.toLowerCase())) {{
            return elem;
        }}
        for (let child of elem.children){{
            let t_elem = recSearchTextExact(child, text, nodeType);
            if (t_elem !== null) {{
                return t_elem;
            }}
        }}
        return null;
    }}
    let elem = document.querySelector('div[class="ui-multiselect-menu ui-widget ui-widget-content ui-corner-all queryConditionSelector statusSelector"]');
    let target = recSearchTextExact(elem, "{text}", "{node_type}")
    
    """
    if toggle_not:
        rec_search_script += """
        elem.querySelector('.notBadge').click()
        """
    # Make final submission
    rec_search_script += """
    target.click()
    """
    return driver.execute_script(rec_search_script)

def areAllFoldersOpen(driver: webdriver.Chrome): 
    script = """
        function isNumeric(num){
            return !isNaN(num)
        }

        function retrieveSum(){
        res = 0
        document.querySelectorAll('#treePane > ul.jstree-children:first-of-type > li > ul > li > a').forEach((e) => {
            let l = e.textContent.split('(')
            n = l[l.length-1].split(')')[0]
            
            if (isNumeric(n)){
                res += Number(n)
            }
            res += 1 // Also count the current + the root at the end
        })
        return res + 1
        }
        let target = retrieveSum()
        function isFullyExpanded(){
            return document.querySelector('ul[role="group"]').querySelectorAll('li[role="treeitem"]').length === target
        }
        return isFullyExpanded()

    """
    return driver.execute_script(script)

def expandAllFolders(driver: webdriver.Chrome):
    print("Expanding all folders... This may take a moment ⌛")
    script = """
        function isNumeric(num){
            return !isNaN(num)
        }
        function recursiveOpen(node, depth){
            if (node === null){
                return
            }
            for (let li of node.querySelectorAll('li[role="treeitem"]')){
                if ( li.classList.contains('jstree-closed')){
                    li.querySelector('i.jstree-icon.jstree-ocl').click()
                }
            }
            node.querySelectorAll('ul[role="group"]').forEach((e)=> {
                recursiveOpen(e, depth + 1)
            })
        }

        // Open the first root
        while ( document.querySelector('ul[role="group"] > li').classList.contains('jstree-closed')){
        document.querySelector('ul[role="group"] > li > i').click()
        }
        function openRoot() {
            while ( document.querySelector('ul[role="group"] > li').classList.contains('jstree-closed')){
                document.querySelector('ul[role="group"] > li > i').click()
            }
        }
        function retrieveSum(){
            res = 0
            document.querySelectorAll('#treePane > ul.jstree-children:first-of-type > li > ul > li > a').forEach((e) => {
                let l = e.textContent.split('(')
                n = l[l.length-1].split(')')[0]
                
                if (isNumeric(n)){
                    res += Number(n)
                }
                res += 1 // Also count the current + the root at the end
            })
            return res + 1
        }
        openRoot()
        let target = retrieveSum()
        const intervalId = setInterval(() => {
        const group = document.querySelector('ul[role="group"]');
        if (document.querySelector('ul[role="group"]').querySelectorAll('li[role="treeitem"]').length !== target) {
            recursiveOpen(group, 0);
        } else {
            console.log("Complete!")
            clearInterval(intervalId);
        }
        }, 3000);


    """
    
    driver.execute_script(script)
        

def select_location(
        driver: webdriver.Chrome, text: str,
        node_type: str, proj_name: str
        ) -> WebElement:
    """ This must be used after using `select_filter()` method and choosing `test location`. This is hacky, but this uses custom selectors based on the project. 

    You must use Double quotes to wrap around your selector!
    """
    if proj_name == 'mib3oigp':
        last_selector = '73303386choiceList11Selector'
    elif proj_name == 'nar classic':
        last_selector = '94014605choiceList11Selector'
    try:
        rec_search_script = f"""
        function recSearchTextExact(elem, text, nodeType) {{
            if ((elem?.nodeName.toLowerCase() === nodeType ) && (elem?.textContent.trim().toLowerCase() === text.toLowerCase())) {{
                return elem;
            }}
            for (let child of elem.children){{
                let t_elem = recSearchTextExact(child, text, nodeType);
                if (t_elem !== null) {{
                    return t_elem;
                }}
            }}
            return null;
        }}
        let elem = document.querySelector('div[class="ui-multiselect-menu ui-widget ui-widget-content ui-corner-all queryConditionSelector {last_selector}"]');
        recSearchTextExact(elem, "{text}", "{node_type}").click();
        """
        return driver.execute_script(rec_search_script)
    except selenium_exceptions.JavascriptException as e:
        print(print(f"The selector for the location filter may have changed. {e}"))
        if DEBUG:
            time.sleep(10000)

def table_search_set_attr(
        driver: webdriver.Chrome, table_selector="#historyList", node_type="tr", text="", attempts=3, 
        attr_key="", attr_val=None
        ) -> tuple[int, WebElement]:
    """ 
        Checks the entry of the table with `text` entry.
        Assigns the attribute of the found element with given value.
        Returns a tuple of the following:
        0: The`ith`row of the table where the text is found
        1: The element which contains the text
    """
    for _ in range(attempts):
        try:
            check_search_res = f"""
    function get_index_tr(test_case_name, css_selector, nodeType, attr, attr_v) {{
        function recSearchTextAttr(elem, text, nodeType) {{
            if (elem?.nodeName.toLowerCase() === nodeType && elem?.textContent.trim().toLowerCase().includes(text.toLowerCase())) {{
                elem.setAttribute(attr, attr_v)
                return elem;
            }}
            for (let child of elem.children) {{
                let t_elem = recSearchTextAttr(child, text, nodeType);
                if (t_elem !== null) {{
                    return t_elem;
                }}
            }}
            return null;
        }}
        let x = document.querySelectorAll(`${{css_selector}} tbody > tr`);
        let i = 0;
        for (let e of x) {{
            let elem = recSearchTextAttr(e, test_case_name, nodeType);
            if (!e.style.display.toLowerCase().includes("none") && elem) {{
                while (!x[i].querySelector('input').checked) {{
                    x[i].querySelector('input').click();
                }}
                return [i, elem];
            }}
            i += 1;
        }}
        return -1;
    }}
    return get_index_tr('{text}', '{table_selector}', '{node_type}' ,'{attr_key}', '{attr_val}');
            """
            return driver.execute_script(check_search_res)
        except selenium_exceptions.StaleElementReferenceException as e:
            print("Element became stale 🍿. Trying again.")
            time.sleep(0.2)
            pass
        

def table_search(
        driver: webdriver.Chrome, table_selector="#historyList", test_set_name="", attempts=3, attr=""
        ) -> tuple[int, WebElement, str] | int:
    """ 
    Returns a list of the following:
    0: The`ith`row of the table where the text is found
    1: The element which contains the text
    2: The attribute of the element that contains the text 

    If none are found, returns -1
    """
    print(f"Searching for entry '{test_set_name}' in the table... ")
    for _ in range(attempts):
        try:
            check_search_res = f"""
    function get_index_tr(test_case_name, css_selector, attr) {{
        function recSearchTextAttr(elem, text, nodeType) {{
            if (elem.nodeName.toLowerCase() === nodeType && elem.textContent.trim().toLowerCase().includes(text.toLowerCase())) {{
                return [elem, elem.getAttribute(attr)];
            }}
            for (let child of elem.children) {{
                let t_elem = recSearchTextAttr(child, text, nodeType);
                if (t_elem !== null) {{
                    return t_elem;
                }}
            }}
            return null;
        }}
        let x = document.querySelectorAll(`${{css_selector}} tbody > tr`);
        let i = 0;
        for (let e of x) {{
            let tup = recSearchTextAttr(e, test_case_name, "a");
            if (!e.style.display.toLowerCase().includes("none") && tup) {{
                while (!x[i].querySelector('input').checked) {{
                    x[i].querySelector('input').click();
                }}
                return [i, tup[0], tup[1]];
            }}
            i += 1;
        }}
        return -1;
    }}
    return get_index_tr('{test_set_name}', '{table_selector}', '{attr}');
            """
            return driver.execute_script(check_search_res)
        except selenium_exceptions.StaleElementReferenceException as e:
            print("Element became stale 🍿. Trying again.")
            time.sleep(0.2)
            pass
        
def search_and_click_on(
        driver: webdriver.Chrome, parent_selector:str,
        textContent: str, elem_type="button",
        tries= 3, timeout = 3
        ) -> None:
    """ Attempts to search the child """
    wait = WebDriverWait(driver, timeout)
    for _ in range(tries):
        try:
            # First, ensure that the parent is found.
            wait.until(
                EC.visibility_of_element_located(
                    (By.CSS_SELECTOR, parent_selector)
                )
            )
           
            # Then search its children for the result.
            wait.until(lambda d: recursive_search_includes(d, parent_selector, textContent, elem_type)).click()
            return

        except selenium_exceptions.StaleElementReferenceException as e:
            print("Element became stale 🍿. Trying again.")
            time.sleep(0.1)
            pass  # Ignore the exception and continue the loop to retry finding the element

        except (selenium_exceptions.TimeoutException, AttributeError, selenium_exceptions.JavascriptException) as e:
            print(f"There's no entry with the text, '{textContent}'. ")
           
def set_attribute(
        driver:webdriver.Chrome, selector:str,
        attr_key: str, attr_val: str
        ) -> None:

    driver.execute_script(
        f"document.querySelector('{selector}').setAttribute('{attr_key}', '{attr_val}')")
    return

    
def wait_till_loading_fin(wait:WebDriverWait):
    """ Waits for the loading popup to be dismissed. """
    try:
        wait.until(EC.visibility_of_element_located(
            (By.CSS_SELECTOR, ".ui-widget-overlay.ui-front"))
        )
        wait.until(
            EC.invisibility_of_element_located(
                (By.CSS_SELECTOR, ".ui-widget-overlay.ui-front")
            )
        )   
    except selenium_exceptions.TimeoutException as e:
        print("Loading banner not shown -- continuing normally...")

def click_on(locator:tuple, driver, timeout=3):
    wait = WebDriverWait(driver, timeout)
    element = None
    for _ in range(timeout):
        try:
            element = wait.until(EC.element_to_be_clickable(locator))
            element.click()
            break
        except selenium_exceptions.StaleElementReferenceException:
            print("Element became stale 🍿. Trying again.")
            pass 

def click_on_highlighted(driver: webdriver.Chrome):
    """ Clicks the TC and verifies that the item is clicked. """
    while True:
        try:
            # Ensure that the entry is clicked.
            click_on(
                (By.CSS_SELECTOR, ".jstree-anchor.jstree-search"), 
                driver, 3)
           
            if "jstree-clicked" in driver.find_element(
                    By.CSS_SELECTOR, ".jstree-anchor.jstree-search"
                ).get_attribute("class"):
                break
        except selenium_exceptions.StaleElementReferenceException as e:
            print("Stale, but trying again.")
            time.sleep(0.2)
            continue
        
def context_click_testcase(
        driver: webdriver.Chrome,  test_case_name:str,
        attempts:int, timeout=3
        ) -> None:
    wait = WebDriverWait(driver,timeout)
    action_chains = AC.ActionChains(driver)
    for _ in range(attempts):
        try:
            action_chains.context_click(recursive_search_includes(driver,'#treePane',test_case_name,"a" )).perform()
            wait.until(
                EC.visibility_of_any_elements_located(
                    (By.CSS_SELECTOR, "ul[class='vakata-context jstree-contextmenu jstree-default-contextmenu']")
                )
            )
            break
        except selenium_exceptions.TimeoutException as e:
            print("Couldn't right-click 😢. Trying again...")
            time.sleep(0.5)
            pass
        
        
def select_from_context_menu(
        driver: webdriver.Chrome, target_text_content: str, timeout=10) -> None:
    """ You must use Double quotes to wrap around your selector! """
    while True:
        try:
            recursive_search_includes(
                driver,
                "ul[class='vakata-context jstree-contextmenu "\
                "jstree-default-contextmenu']", target_text_content, "a" 
            ).click()
            break
        except selenium_exceptions.StaleElementReferenceException as e:
            print("Stale element. Trying again.")
            time.sleep(0.5)
            continue
        except AttributeError:
            print("Doesn't seem to detect context menu options... Trying again")
            time.sleep(0.5)
        except selenium_exceptions.TimeoutException as e:
            print("I couldn't open up the context menu 😢. "\
                  "Try not to move the mouse around.")
            raise selenium_exceptions.TimeoutException
            

def select_from_dropdown_menu(
        driver: webdriver.Chrome, selector:str,
        target_text_content: str, elem_tag = "option" ):
    """ You must use Double quotes to wrap around your selector! 
        Customized on a project basis.
    
    """

    while True:
        try:
            recursive_search_includes(driver, selector, target_text_content, elem_tag).click()
            break
        except selenium_exceptions.StaleElementReferenceException as e:
            print("Stale element. Trying again.")
            continue

def handle_child_tc_warning(driver: webdriver.Chrome, parent_css_selector="div.ui-dialog-buttonset") -> bool:
    """ Looks for the warning, then switches to the iframe. """
    try:
        # Click on the option
        search_and_click_on(driver, parent_css_selector , "selected items")
        print("Child TC found. Selecting only chosen TC's only.")

        # Then switch to iframe
        switch_to_iframe(driver)
        return True
    except selenium_exceptions.TimeoutException:
        print("Couldn't locate the warning button 😢")
        return False

def switch_to_iframe(
        driver:webdriver.Chrome, timeout=4, css_selector="#inlinedPopupIframe", tries=3
        ) -> bool:
    """ Switches to the popup iframe, if any. """
    wait = WebDriverWait(driver, timeout)
    for i in range(tries):
        try: 
            """ Waits first before switching to the iframe. Returns true once switched """
            driver.switch_to.frame(wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, css_selector))))
            return True
        except selenium_exceptions.TimeoutException as e:
            print("Waiting for the popup...⌛")
            time.sleep(0.2)
            pass
    return False

def expand_all_folders(driver:webdriver.Chrome):
    script = """
        function recursiveOpen(node, depth){
		    if (node == null){
		        return
		    }
		    for (let li of node.querySelectorAll('li[role="treeitem"]')){
		        if (depth === 0 || li.classList.contains('jstree-closed')){
		            li.querySelector('i.jstree-icon.jstree-ocl').click()
		        }
		    }
		    node.querySelectorAll('ul[role="group"]').forEach((e)=> {
		        recursiveOpen(e, depth + 1)
		    })
		}
		recursiveOpen(document.querySelector('ul[role="group"]'), 0)
    """

def create_test_set(driver: webdriver.Chrome, df: pd.DataFrame, component_name:str, res_col_name:str) -> str:
    """ Creates a test set based on everything on the existing 'name' column
        on the spreadsheet.  
        `Component_name` and `res_col_name` is needed for naming purposes.
        the `res_col_name` variable should be formatted in the format:
        *  "`<version> <priorities, if any>`"
        Returns a link to the test set
    """
    if DEBUG:
        print("Creation of a test set has started. Please do not move your mouse during the process! ✋")
    # If the test set already exists, use that one.
    wait = WebDriverWait(driver, 20)
    short_wait = WebDriverWait(driver, 2)
    action_chains = AC.ActionChains(driver)
    
    # Use this test set name to search and add.
    ver_number = res_col_name[:7].upper()
    t_desc = res_col_name[7:].strip().title()
    ver_desc = f" | {t_desc}" if len(t_desc) > 1 else ''
    test_set_name = f"[{component_name.upper()}][{ver_number}] Test Set{ver_desc}"
     
    wait.until(EC.visibility_of_element_located((By.ID, "searchBox_treePane")))
    wait.until(EC.visibility_of_element_located((By.ID, "go_treePane")))
    wait.until(EC.visibility_of_element_located(
        (By.CSS_SELECTOR, ".ui-dialog-content")))
    wait.until(EC.invisibility_of_element_located(
        (By.CSS_SELECTOR, ".ui-dialog-content")))
   
    # Remove any existing filters, if any
    try:
        [
            x.click() for x in wait.until(
                EC.visibility_of_all_elements_located((By.CSS_SELECTOR, '.modifiableArea > span.ui-icon.removeIcon'))
            )
        ]
    except selenium_exceptions.TimeoutException as e:
        pass

    # Apply 2 filters
    # Open filter window
    click_on((By.CSS_SELECTOR, 'span.filterLabel'), driver, 20)
    
    # Choose test location
    # Window 1
    select_filter(driver, "test location", "span")
    select_location(driver, 'NAR', 'span', proj_name)

    # Open filter window
    click_on((By.CSS_SELECTOR, 'span.filterLabel'), driver, 20)
    
    wait.until(EC.presence_of_element_located(
        (By.CSS_SELECTOR, 
        '.ui-multiselect-menu > ul[class="ui-multiselect-checkboxes ui-helper-reset')))

    select_filter(driver, "status", "span")
    select_status(driver, "deprecated", "span", True)

    # Click GO to apply the filter
    click_on((By.CSS_SELECTOR, "#actionBarSearchButton"), driver, 3)

    # Wait until loading is finished
    wait_till_loading_fin(wait)     
   
    # Expand all folder
    expandAllFolders(driver)
    
    # Wait until the folders are open
    WebDriverWait(driver,250,1).until(lambda d: areAllFoldersOpen(d))

    def tc_is_selected(driver: webdriver.Chrome, elem:WebElement ):
        if not elem:
            raise AttributeError
        return "jstree-clicked" in elem.get_attribute('class')
    i = 0

    # To 
    last_successful_tc = ''
    for df_row in df.iterrows():
        for tries in range(2, -1, -1):
            try:
                # Switch to default iframe
                driver.switch_to.default_content()
                
                tc_elem = driver.find_element(By.CSS_SELECTOR, f"""li[id="{df_row[1]['id']}"] > a""")

                # Only click the first item, then hold ctrl thereafter
                if i == 0:
                    action_chains.click(tc_elem).perform()
                else:
                    action_chains.key_down(Keys.CONTROL).click(tc_elem).key_up(Keys.CONTROL).perform()

                short_wait.until(lambda d: tc_is_selected(d, tc_elem))
                last_successful_tc = df_row[1]['name']
                break
            except selenium_exceptions.TimeoutException as e:
                action_chains.reset_actions()
                print(f"Timed out for this test case: {df_row[1]['name']}. "\
                      "Trying {tries} more times...")
                pass
            except selenium_exceptions.NoSuchElementException as e:
                action_chains.reset_actions()
                print(f"I could not find the test case: {df_row[1]['name']} 😢")
                break
        i += 1
    # At the final test case, right click, if possible.
    if last_successful_tc:
        context_click_testcase(driver, last_successful_tc, 3)
    else:
        raise Exception("No test cases were selected. "\
                        "If this was a mistake, run the script again. ")
    # Wait until the context menu appears and select
    # Try to add to an existing test set
    select_from_context_menu(
        driver, "Add selected Test Cases to Test Set...", 10
    )

    # Handle if there's a selected child TC, select "add selected"
    wait.until(
       lambda d: switch_to_iframe(d) or handle_child_tc_warning(d)
       )
            
    # Expecting an iframe here.
    # Now, switch to the 'inlinedPopupIframe' by its name
    # switch_to_iframe(driver, wait, "#inlinedPopupIframe")
    try:
        click_on((By.CSS_SELECTOR, "#historyTab-tab"),driver, 3)

        driver.execute_script(
            """document.querySelector("#filterInput").removeAttribute('maxLength')""")
        click_on((By.CSS_SELECTOR, "#filterInput"), driver, 3)
        driver.find_element(By.CSS_SELECTOR, "#filterInput").send_keys(test_set_name)

        idx = table_search(driver, test_set_name=test_set_name)
    except selenium_exceptions.TimeoutException as e:
        # Try to search using "Find Test Set" column
        click_on((By.CSS_SELECTOR, "#searchTab-tab"), driver, 3 )
        
        # type in search pattern
        driver.find_element(By.CSS_SELECTOR, '#searchPattern').send_keys(test_set_name)

        # Press search
        click_on((By.CSS_SELECTOR, '#searchButton'), driver, 3)

        # Wait until results show
        time.sleep(2)

        idx = table_search(driver, "#searchList", test_set_name)[0]


    # The test set is not found
    if idx == -1:  
        print("There's no existing test set for this. Creating a new one for you...")
        click_on((By.CSS_SELECTOR,'#command .cancelButton'), driver, 3)

        # Switch to default iframe
        driver.switch_to.default_content()
        # click_on_highlighted(driver) 

        # final test case, right click 
        action_chains.context_click(recursive_search_includes(driver,'#treePane',last_successful_tc,"a" )).perform()

        # Wait until the context menu appears
        wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "ul[class='vakata-context jstree-contextmenu jstree-default-contextmenu']")))

        # Select Create Test set.
        select_from_context_menu(driver, "Create Test Set from selected Test Cases...")

        # Expecting an iframe here once again.
        # Handle if there's a selected child TC, select "add selected"
        wait.until(
            lambda d: switch_to_iframe(d) or handle_child_tc_warning(d)
        )

        # Configure Tracker
        wait.until(
            EC.visibility_of_element_located(
                (By.CSS_SELECTOR, 'select[name="tracker_id"]')
                )
            )

        select_from_dropdown_menu(
            driver,"select[name='tracker_id']", 
            CONFIG["settings"][proj_name]["test_set_tracker"])

        # Configure project
        select_from_dropdown_menu(driver, "#project", "VW Cockpit 2022+")

        click_on((By.CSS_SELECTOR, 'input[name="createNewTestSet"]'),driver, 3)

        wait.until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, "#summary"))
            ).send_keys(test_set_name)

        click_on((By.CSS_SELECTOR, 'input[name="SUBMIT"]'), driver, 3)

        wait.until(
            EC.visibility_of_element_located(
                (By.CSS_SELECTOR, 'input[type="submit"]'))
            )

        # Save the link before closing 
        href_link = recursive_search_incl_get_attr(driver, '.contentArea', test_set_name, 'a', 'href')

        click_on((By.CSS_SELECTOR,'input[type="submit"]'), driver, 3)

        wait.until(EC.invisibility_of_element_located((By.CSS_SELECTOR, ".ui-widget-overlay.ui-front")))
        pass
    else:
        print("Found an existing test set! Appending onto the existing set 😎")
        
        # Add to existing set
        click_on((By.CSS_SELECTOR, 'input[name="addToExistingSet"]'), driver, 3)

        # Close the dialogue button (Should also handle duplicate as well)
        wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, 'input[value="Close"]')))

        # Save the link before closing 
        href_link = recursive_search_incl_get_attr(driver, '.contentArea', test_set_name, 'a', 'href')

        click_on((By.CSS_SELECTOR, 'input[value="Close"]'), driver, 3)
        # Deselect the result
    
    print("Adding complete!")
    # Save the URL
    return "http://vwavncb.lge.com:8080" + href_link[1]



def url_extractor(s: str):
    """ Extracts codeBeamer links from a plain text """
    start = s.find('http://')
    urls = []
    while start >= 0 :
        end = s.find('\n', start)
        if end == -1: end = s.find(' ', start)
        if end == -1: 
            end = find_nth(s, '/', 5, start) + 1 # Get 1 after the slash
            if end != -1 and end+8 <= len(s) and s[end:end+8].isnumeric():
                # Not a valid CB url.
                end = end+8
            else:
                end = -1
        if end != -1:
            r_url = s[start:end].split('\n')[0].strip()
            urls.append(r_url)
            start = s.find('http://', start + len(r_url))
        else:
            start = s.find(' ', start)
    return urls
       
def find_nth(haystack:str, needle: str, n: int, offset = 0) -> int:
    """ Who knew a LC problem would come in handy? """
    start = haystack.find(needle, offset)
    while start >= 0 and n > 1:
        start = haystack.find(needle, start + len(needle))
        n -= 1
    return start

def cb_login(driver: webdriver.Chrome, proj_name:str, component_name: str):
    """ Navigates to the specified CodeBeamer link based on `proj_name` 
    performs login. The `component_name` isn't used yet. Probably won't need it.
    precondition: The environment variables and `CONFIG` must have 
    been loaded already.
    """
    link = CONFIG["settings"][proj_name]["components"][component_name]\
        ["test_case_link"]
    try:
        driver.get(link)
        driver.find_element(By.ID, "user").send_keys(CB_ID)
        driver.find_element(By.ID, "password").send_keys(CB_PASS)
        driver.find_element(
            By.CSS_SELECTOR, value="input[value='Login']").click()
       
    except IncorrectLoginCredentials as e:
        raise Exception(e)
    except selenium_exceptions.NoSuchElementException as e:
        raise CodeBeamerMaintenance
        
def verify_if_correct_test_case(
        wait: WebDriverWait, spreadsheet_row_n:int, test_set_name:str):
    try:
        testset_count = wait.until(EC.visibility_of_element_located(
            (By.CSS_SELECTOR, "#testSetTestCases-tab"))
            ).text
        testrun_summary = wait.until(EC.visibility_of_element_located(
            (By.CSS_SELECTOR, ".breadcrumbs-summary>a.generated-link"))
            ).text
    except selenium_exceptions.TimeoutException as e:
        raise selenium_exceptions.WebDriverException(
            "I could not determine the number of test cases from this page. "\
            "\nDid you provide the correct link to the test set? 😗")
    testset_count = str_to_int(
        testset_count.strip().split('Test Cases & Sets (')[1].split(')')[0]
    )
    
    if testset_count != spreadsheet_row_n:
        raise IncompleteColumnError(
            f"There are mismatching number of test cases! "\
            "Spreadsheet count: {spreadsheet_row_n} | Test set count: "\
            "{testset_count}")
    if testrun_summary.strip().lower() != test_set_name.lower():
        raise IncompleteColumnError(
            f"This doesn't appear to be the expected test set: "\
            "'{test_set_name}'. Ensure that you are using a created a test "\
            "set with the script beforehand.")
        

def do_test_run(
        driver: webdriver.Chrome, df: DataFrame, res_col_name:str, component_name: str, proj_name:str, test_set_link:str):
    """ Main driver to perform a test run. DataFrame is assumed to be 
        filtered and is following the requirements specified in the guideline.
    """
    
    ver_number = res_col_name[:7].upper()
    t_desc = res_col_name[7:].strip().title()
    ver_desc = f" | {t_desc}" if len(t_desc) > 1 else ''
    test_set_name = f"[{component_name.upper()}][{ver_number}] Test Set{ver_desc}"
    test_run_name = f"[{component_name.upper()}][{ver_number}] Test Run{ver_desc.title()}"
    test_run_item = CONFIG["settings"][proj_name]["test_run_item_prefix"] \
                    + ver_number.split('.')[0]
    test_config:str = CONFIG["settings"][proj_name]["test_configuration"]
    spreadsheet_row_n = df.shape[0]

    wait = WebDriverWait(driver, 20)

    if DEBUG:
        input("Creating a test run will now begin. "\
              "\nPress 'Enter' to continue. Otherwise, press ctrl + c: ")

    if test_set_link:
        driver.get(test_set_link)
    else:
        driver.get(CONFIG["settings"][proj_name]["components"][component_name]["test_set_link"])

        # Try to wait for the loading banner before searching
        wait_till_loading_fin(wait)

        search_elem = wait.until(
            EC.element_to_be_clickable(
                (By.CSS_SELECTOR, 
                 'input[title="Full text search for all fields including '\
                'comments and attachments."]')
                )
            )

        # Search on the current table if it shows.
        res_tup = table_search(driver, "#trackerItems", test_set_name,3, "data-id" )

        # If not found, try to apply the filter
        if res_tup == -1:
            click_on(
                (By.CSS_SELECTOR, 
                'input[title="Full text search for all fields including '\
                'comments and attachments."]'), driver)
            search_elem.send_keys(Keys.CONTROL + "a")
            search_elem.send_keys(Keys.DELETE)
            search_elem.send_keys(test_set_name)

            # Press "GO"
            click_on((By.CSS_SELECTOR, "#actionBarSearchButton"), driver)

            # Wait until loading is finished
            wait_till_loading_fin(wait)

            # Search table
            res_tup = table_search(
                driver, "#trackerItems", test_set_name,3, "data-id" 
            )

        # If still not found, then allow the user to input the tc link 
        # manually or quit.
        if res_tup == -1:
            while True:
                print("No test set is found. If you recently created this "\
                      "test set, it will take some time for it to appear on "\
                      "the codeBeamer DB.\nFor now, you need to provide a "\
                      "valid test set link.")
                test_set_link = input("Enter a valid test set URL: ")
                if valid_test_link(test_set_link):
                    driver.get(test_set_link)
                    break
        else:        
            click_on((By.CSS_SELECTOR, f"a[data-id='{res_tup[2]}']"), driver, 20)    
    
    # Check if there are the same number of test cases
    verify_if_correct_test_case(wait, spreadsheet_row_n, test_set_name)

    # Press the play icon
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'img[title="New Test Run"'))).click()
    
    # Search and click item
    search_and_click_on(driver, "#ui-id-5", test_run_item , 'a')

    # Check the two checkmarks
    wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, '#formalityRegularLabel > input')))
    driver.execute_script(
    """
        document.querySelector('#formalityRegularLabel > input')?.setAttribute('checked', true)
        document.querySelector('#runOnlyAcceptedTestCases2')?.setAttribute('checked', true)
        document.querySelector('input[id="testrun.editor.distributeRunsStrategy.sharedBetweenMembers"]')?.setAttribute('checked', true) 
    """)

    # Select NAR
    recursive_search_exact_set_attr(driver,'#dynamicChoice_references_1000', 'NAR')

    # Write name
    summary_elem = driver.find_element(By.CSS_SELECTOR,"#summary")
    summary_elem.send_keys(Keys.CONTROL + "a")
    summary_elem.send_keys(Keys.DELETE)
    summary_elem.send_keys(test_run_name)
    
    # Select releaseID
    recursive_search_exact_set_attr(driver,'#releaseId', ver_number.split('.')[0] )

    # Select config
    e_tuple = recursive_search_incl_get_attr(driver, 
    '#testRunConfigurationsList', test_config, 'label' , 'for')
    click_on((By.CSS_SELECTOR, f"#{e_tuple[1]}"), driver )

    # Hit save
    click_on((By.CSS_SELECTOR, 'input[title="Save (Ctrl + S)"]'), driver)
    
    if DEBUG:
        input(f"The test run will now begin for the test set: {test_set_name}. Press 'Enter' to continue. Otherwise, press ctrl + c.")
        
    # Up to this point, the test sets are ready!
    try:
        wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR,'.actionBar a[title="Run!"]')))
        click_on((By.CSS_SELECTOR, '.actionBar a[title="Run!"]'), driver)

        # Switch to the new tab
        curr_window = driver.current_window_handle
        
        wait.until(lambda d: len(d.window_handles) == 2)

        for window in driver.window_handles:
            if window != curr_window:
                driver.switch_to.window(window)
                driver.maximize_window()
                print("switched to the new window.")
    
        # Store the number of test cases total
        tc_metadata_elem:WebElement = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, 'span[title="Number of Tests finished."]')))

        # Extract the counts
        tc_str = tc_metadata_elem.text.strip().split('of')
    
        no_finished = str_to_int(tc_str[0].strip())
        total_tc = str_to_int(tc_str[1].strip().split(' ')[0].strip())
        if no_finished == None or total_tc == None:
            raise ValueError(f"Could not retrieve the metadata of the tc's: {tc_str}")
        if total_tc != spreadsheet_row_n:
            print(f"There are mismatching number of tc's on this test run and your tc's on your spreadsheet. Test run: {total_tc} | your spreadsheet: {spreadsheet_row_n}")
    except ValueError as e:
        print(e)
        no_finished = 0
        total_tc = spreadsheet_row_n # Total number of rows
    if DEBUG:
        # print(f"range: {no_finished} | {total_tc}")
        print("Test run has started. Please do not move your mouse during the process! ✋")

    def get_tc_name(driver: webdriver.Chrome):
        tc_name = driver.find_element(By.CSS_SELECTOR, '#summaryTd a:last-child').get_attribute("title")
        begin = tc_name.find("]") + 2
        tc_name = tc_name[begin:].strip().lower()
        return tc_name
    
    def get_tc_id(driver: webdriver.Chrome):
        """ The href attribute is always expected to have the format of:
            `/cb/item/<cb_id>?<some_queries>`
         """
        tc_href = driver.find_element(By.CSS_SELECTOR, '#summaryTd a:last-child').get_attribute("href")
        return tc_href.split('/cb/item/')[1].split('?')[0]


    def get_current_turn(driver: webdriver.Chrome):
        try:
            return WebDriverWait(driver,5).until(EC.visibility_of_element_located((By.CSS_SELECTOR, '#jumpTo'))).get_attribute('value')
        except AttributeError as e:
            raise selenium_exceptions.NoSuchWindowException
    current_turn = ''
    for i in range(no_finished, total_tc):
        # Wait until the buttons are interactable
        try:
            wait.until(lambda d: current_turn != get_current_turn(d))
            current_turn = get_current_turn(driver)

            # Get the tc name
            tc_name = get_tc_name(driver)
            tc_id = get_tc_id(driver)

            # Filter and get the row by name
            row_df = df.loc[df["id"] == tc_id]
            if row_df.size == 0:
                raise NoEntryFound(f"Row with {tc_name} with id {tc_id} is not found on your spreadsheet 😪")
            row_df = row_df.iloc[0]
            result = row_df[res_col_name]
            comment = row_df["comments"]
            
            # The test case passes
            if result == 'pass':
                while True:
                    try:
                        click_on((By.CSS_SELECTOR, '#buttonTable tr > td button:first-of-type'), driver, 10)
                        
                        driver.find_element(By.CSS_SELECTOR, '#conclusionInDialog')
                        break
                    except selenium_exceptions.NoSuchElementException as e:
                        time.sleep(0.5)
                        continue

            elif result == 'fail':
                # Report all bugs first
                cb_link: list[str] = url_extractor(comment)
                if len(cb_link) == 0:
                    print(f"You need to include a cb ticket link for this test case: {tc_name}")

                for link in cb_link:
                    click_on((By.CSS_SELECTOR, '#reportBugButton'), driver, 10)

                    # Switch to separate iframe
                    switch_to_iframe(driver, 10)

                    click_on((By.CSS_SELECTOR, "#findAnExistingBug-tab"), driver,10)
                    start = link.rfind('/')+1
                    cb_code = link[start:start+8] # 8 digit code
                    wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, '#searchForBug'))).send_keys(cb_code)

                    # Wait until the search results to show
                    wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "#result > div")))

                    table_search_set_attr(driver, "#result>table", "tr",cb_code, attr_key="checked", attr_val="true")

                    click_on((By.CSS_SELECTOR, '#findAnExistingBug input[value="Add selected Bugs"]'), driver)

                    driver.switch_to.default_content()

                # Then proceed to submit failed test case
                click_on((By.CSS_SELECTOR, 'button[name="failStep"]'), driver, 10)

                # Write the comment
                wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, '#conclusionInDialog'))).send_keys(comment)

                pass
            elif result == 'blocked' or result == 'na':
                click_on((By.CSS_SELECTOR, 'button[name="blockStep"]'), driver, 10)
                
                wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, '#conclusionInDialog'))).send_keys(comment)

            # Input the summary 
            search_and_click_on(driver, 'body > div.ui-dialog.ui-corner-all.ui-widget.ui-widget-content.ui-front.cbModalDialog.ui-dialog-buttons.ui-resizable > div.ui-dialog-buttonpane.ui-widget-content.ui-helper-clearfix > div', 'save')
            
        except selenium_exceptions.TimeoutException as e:
            print(f"Something went wrong. TC: '{tc_name}', {e}")
            if i < total_tc-1:
                click_on((By.CSS_SELECTOR, '#next'), driver, 10)
                WebDriverWait(driver, 20, 0.5).until(lambda d: get_tc_name(d) != tc_name)
            continue
        except (
            selenium_exceptions.NoSuchElementException, 
            selenium_exceptions.StaleElementReferenceException) as e:
            raise selenium_exceptions.WebDriverException("Something went wrong during the test run process.")
        except NoEntryFound as e:
            print(e)
            if i < total_tc-1:
                click_on((By.CSS_SELECTOR, '#next'), driver, 10)
                WebDriverWait(driver, 20, 0.5).until(lambda d: get_tc_name(d) != tc_name)
            continue
        except selenium_exceptions.NoSuchWindowException as e:
            print("The window has been closed.")
            break

def print_credits(name, email):
    title = "codeBeamer ALM Test Suit Pilot"
    subtitle = "Create and run tests! Release - v1.0"
    name_field = f"Credits: {name} | License: MIT"
    email_field = f"Bug reports: In person or email to {email}"
    
    max_length = max(len(title), len(subtitle), len(name_field), len(email_field)) + 4  # Add some padding
    
    border = '*' * (max_length + 2 )
    
    credits = f"""
    {border}
    * {title.center(max_length-2)} *
    * {subtitle.center(max_length-2)} *
    {border}
    * {name_field.center(max_length-2)} *
    * {email_field.center(max_length-2)} *
    {border}
    """
    print(credits)

# Main script
if __name__ == "__main__":
    try:
        # Start scraping!
        driver = configure_webdriver()
        global CONFIG
        load_env()
        CONFIG = load_config()
        update_config(CONFIG)
        test_set_link = ''

        print_credits("SeungJoon Yang", "tmdwns.yang@gmail.com")
        
        # Get user input to select which project
        proj_name, component_name, create_test_set_ans = get_user_selection(CONFIG)

        # Read get the spreadsheet 
        df = get_excel(proj_name, component_name)

        # Perform clean up
        # Verify if the results column and the name columns are valid
        df, res_col_name = setup_df(df, proj_name, component_name)
        
        # Perform login
        cb_login(driver, proj_name, component_name)
        
        # Start creating/adding test cases
        if create_test_set_ans:
            test_set_link = create_test_set(driver, df, component_name, res_col_name)
        
        # Perform test run
        do_test_run(driver, df, res_col_name, component_name, proj_name, test_set_link)

        print("Successful run! 👏 Nice work!")


    except KeyboardInterrupt as e:
        print("Aborted by user.")
    except FileNotFoundError as e:
        print(e)
    except IncompleteColumnError as e:
        print(e)
    except selenium_exceptions.WebDriverException as e:
        msg = str(e)
        start = msg.find("Message")
        end = msg[start:].find('\n')
        print(msg[start:end])
        # print(msg)
    except CodeBeamerMaintenance as e:
        print(e)
    except Exception as e:
        print(e)
    finally:
        input("Press 'Enter' key to exit...")
        print("Exiting...")
        print("Cleaning up 🧹... Please wait")
        driver.quit()
    