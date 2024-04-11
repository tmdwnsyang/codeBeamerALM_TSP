<!---
Copyright 2024 SeungJoon Yang. All rights reserved.

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
-->
## codeBeamer ALM Test Suit Pilot
**Description**: An automation tool for codeBeamer ALM to create, manage, and run bulk test runs with only a few clicks. For personal use only. This project is not affiliated with LGE in any shape or form.

**Author**: SeungJoon Yang

## Features
-   Automate the creation of bulk test sets and perform test runs with inputted test data on the spreadsheet.
- Perform concurrent test runs for different projects and components on a single machine.
-   Intelligently auto-detects the results and filter the test cases based on corresponding results, comments, and codeBeamer tickets provided in the spreadsheet.
-   Ensure compatibility with all currently pending projects and components, and future projects that utilize codeBeamer ALM.
-   Ensure the validity of all test results by checking for inconsistencies, errors, or duplicates to minimize errors.
-  Auto detect ticket links and embeds the corresponding tickets on codeBeamer for failed test cases during the test runs.
-   Configure test set presets and test run settings automatically on a project basis.

 


## 1. Usage
### I. Setting up the `.env` file
The script needs your codeBeamer credentials for authorization, as well as the directory of your test result reports, organized by project and components. In the provided `.env` file, you must set the values for the following: 
* `SYQT_HOME_DIR` - Absolute path containing all of your test reports, organized by project and component basis. By default, the home directory is set to `C:\Users\lgeuser\Documents`
* `CB_ID` - codeBeamer ID 
* `CB_PASS` - codeBeamer password

### II. Expectations
The script primarily performs three sequential steps:
1. Creating a test set based on the project selected
2. Preparing the test run
3. Performing the test run

Before initiating the first step, the script ensures that the spreadsheet is valid. To guarantee the script’s successful execution, please meet all the requirements specified in the following section, [Setting Up Your Spreadsheet](#iii.-setting-up-your-spreadsheet). 

In the first step, the script verifies whether an existing test set is present by searching the codeBeamer database.* In an effort to prevent duplicate entries, newly selected test cases are appended to the existing set*. Otherwise, a new test set is created. Once the test set is successfully created, it can be accessed via codeBeamer. The newly created test set, named `[component name][version number] Test Set [priority description]`, can be found in the corresponding project directories. Currently, the supported projects are NAR Classic and MIB3GP. Support for ICAS3GP is under consideration for future releases.

In the second step, prior to performing the test runs, the following configuration options are set for the NAR:

-   Formality and distribution of the test run among members
-   Test location
-   Name
-   Release
-   Test Configuration

In the third step, the test run is performed as usual. For failed test cases, the script detects any corresponding codeBeamer ticket links contained in the `Comments` column of the spreadsheet and embeds them in the results. For blocked test cases, the associated comments are added.

### III. Setting Up Your Spreadsheet
-   Your spreadsheet name must adhere to the naming convention:
    -   `[<Project Name>][<Component>] SyQT Test Case Full.xlsx`
-   Your spreadsheet should include the following columns:
    -   `Comments`
        -   For a failed test case, one or more KPM links must be included.
    -   `Name`
        -   Identifier of the test cases (duplicates are accepted).
    - `id`
	    - Unique ID for each test case
    -   `<Results Column>`
        -   This column must contain one or more of the following values:  `pass`,  `fail`,  `blocked`,  `na`,  `excl`.
        -   The column name should start with a full version alias, followed by a description, such as:
            -   `NXXX.XX P0 + P1`
        -   Use the  `excl`  value to denote the exclusion of test cases that are not required. Test cases with this value will not be included in the test set.
-   Your spreadsheet must have an anchor column that corresponds to the total number of results.
-  Any duplicate test case results must have identical comments and results.
- **NOTE**: If any of the rules are violated, the script will output a warning with the details and exit.  

### IV. Configuration File (optional)

The configuration file sets up each project component. Below is an example for a NAV component in NAR Classic:

```json
"nav": {
    "anchor_column": "Name",
    "test_case_link": "http://vwavncb.lge.com:8080/cb/tracker/73303386",
    "test_set_link": "http://vwavncb.lge.com:8080/cb/category/77502497"
}
```
**Customizing values for the configuration file**

-   The  `anchor_column`  verifies the match between the number of test cases and entries in the  `results`  column. Ensure the  `anchor_column`  has corresponding non-empty values in the  `results`  column.
-   `test_case_link`  should point to the smallest possible parent tree in codeBeamer containing relevant test cases to avoid duplicates. Obtain the URL by:
    1.  Selecting the folder with your test cases under “trackers”.
    2.  Right-clicking the folder and choosing “Show tree from this item”.
    3.  Copying the URL into the  `test_case_link`  value.
-   `test_set_link`  is the codeBeamer URL for all test sets, enabling the script to find and execute test runs.

## 2. FAQ
- **Q: What if a test case on the spreadsheet is not found in codeBeamer?**
  - A: The script creates or appends a test set without the test case(s). However, the script will not perform the test run on this set, as it cannot verify the validity of the test set. You can simply exclude the test case using the `excl` keyword in the results column or specify the link to the test case tree that contains all the test cases, in addition to the missing test case(s) in the configuration file specified [here](iv.-configuration-file-(optional)).
  
 - **Q: Are duplicate test cases accepted?**
    - Yes. For example, duplicate test cases that share the same name but have different priority values are accepted.  

 - **Q: Is this be compatible with other testing regions?**
    - No. However, if the region follows a similar methodology in record-keeping for system tests, region-specific test creation and test run configuration can be implemented. If desired, please contact me.

## 3. Bug Reports
If you are running into any issues, please send any bug reports to tmdwns.yang@gmail.com. Please include the following items in your report:
* Your completed spreadsheet
* Description of the issue in detail
* Project or model
