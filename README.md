# new_property
A Python utility to help on-board a new Hotel Property.
<kbd>
  <img src="https://github.com/xmatters/xMatters-Labs/raw/master/media/disclaimer.png">
</kbd>

# Pre-Requisites
* [Python 3.7.1](https://www.python.org/downloads/release/python-371/) (I recommend using [pyenv](https://github.com/pyenv/pyenv) to get and manage your python installations)
* Python [requests](http://docs.python-requests.org/en/master/) module (`pip install requests`)
* Python [openpyxl](https://openpyxl.readthedocs.io/en/stable/) module (`pip install openpyxl`)
* Details for the target xMatters instance (Non-Production, Production, or both)

# Files
* [new_property.py](new_property.py) - Main driver/starting point.
* [config.py](config.py) - Defines the config object used by the program, and error messages
* [np_logger.py](np_logger.py) - Provides logging capabilities to the utility.
* [cli.py](cli.py) - The Command Line processor that handles dealing with command line arguments, as well as rading the defaults.json file.
* [processor.py](processor.py) - The guts of the utility where all of the interactions between the .xlsx file and xMatters occurs
* [defaults.json](defaults.json) - Example default property settings.  You may override these with command line arguments too.

# How it works
The user provides an input spreadsheet (template included) that defines a set of xMatters Sites, Administrative Users (per Site), and a Security Group (per Site).  The utility then reads that information and creates the related objects in either Non-Production or Production instances, and then updates the input Spreadsheet with the UUIDs of the created objects.

# Installation

## Python / pyenv setup
* [Python 3.7.1](https://www.python.org/downloads/release/python-371/) (I recommend using [pyenv](https://github.com/pyenv/pyenv) to get and manage your python installations)
* After pyenv is installed, go to the directory where you downloaded the files from this repository.
   1. `pyenv install 3.7.1` - Installs Python v3.7.1 into your local system (You only need to do this once)
   2. `pyenv local 3.7.1` - Makes Python v3.7.1 the default for _this_ project (you only need to do this once per project)
   3. `eval "$(pyenv init -)"` - [For Linux/Max OS X] Initializes pyenv for command line access to Python  (you need to do this whenver you open a Terminal windows unless you do step 5 below)
   4. `pyenv shell 3.7.1` - Sets up access to Pythons tools from the command line (e.g. `pip`)
   5. `echo 'eval "$(pyenv init -)"' >> ~/.bash_profile` - [For Linux/Mac OS X] - makes sure Pyenv is initialized whenver you open a Terminal window.
   6. `pyenv global 3.7.1` - Sets Python v3.7.1 as your global instance of Python (Optional)
* Install the Python [requests](http://docs.python-requests.org/en/master/) module
   * `pip install requests`
* Install the Python [openpyxl](https://openpyxl.readthedocs.io/en/stable/) module
   * `pip install openpyxl`

## new_property.py setup
All you need to do now is to create an appropriate defaults.json.  Use the included version for an example.


# Running
`Run one of these commands:
* `python3 new_property.py -v -c -d defaults.json sites`
   * Processes Sites only
* `python3 new_property.py -v -c -d defaults.json admins`
   * Processes Admin Users only
* `python3 new_property.py -v -c -d defaults.json groups`
   * Processes Security Groups only
* `python3 new_property.py -v -c -d defaults.json all`
   * Processes Sites, Admin Users, and Security Groups
`   

# Usage / Troubleshooting
```
python3 new_property.py -h

usage: new_property.py [-h] [-c] [-d DEFAULTS_FILENAME]
[-f PROPERTIES_FILENAME] [-i {np,prod}]
[-l LOG_FILENAME] [-o OUT_DIRECTORY] [-p [PASSWORD]]
[-s SUPERVISORS] [-U UDF_NAME] [-u USER] [-V] [-v]
[-x XMOD_URL]
{sites,admins,groups,all} ...

Created by jolin@xmatters.com on 2018-11-18.
Copyright 2018 xmatters, Inc. All rights reserved.

Licensed under the Apache License 2.0
http://www.apache.org/licenses/LICENSE-2.0

Distributed on an "AS IS" basis without warranties
or conditions of any kind, either express or implied.

USAGE

positional arguments:
{sites,admins,groups,all}
sites               Use this command in order to only read and process Sites.
admins              Use this command in order to only read and process Admins.
groups              Use this command in order to only process Groups.
all                 Use this command in order to process all worksheets from the infput file: Sites, Admins, Groups.

optional arguments:

-h, --help            show this help message and exit

-c, --console         If specified, will echo all log output to the console at the requested verbosity based on the -v option

-d DEFAULTS_FILENAME, --defaults DEFAULTS_FILENAME
Specifes the name of the file containing default settings [default: defaults.json]

-f PROPERTIES_FILENAME, --pfile PROPERTIES_FILENAME
If not specified in the defaults file, use this for the input file .xlsx file. [default: None]

-i {np,prod}, --itype {np,prod}
Specifies whether we are updating the Production (prod) or Non-Production (np) instance. [default: np]

-l LOG_FILENAME, --lfile LOG_FILENAME
If not specified in the defaults file, use -l to specify the base name of the log file. The name will have a timestamp and .log appended to the end.

-o OUT_DIRECTORY, --odir OUT_DIRECTORY
If not specified in the defaults file, use -o to specify the file system location where the output files will be written.

-p [PASSWORD]         If not specified in the defaults file, use -p to specify a password either on the command line, or be prompted

-s SUPERVISORS, --supervisors SUPERVISORS
If not specified in the defaults file, use this for the xMatters User IDs of the default Supervisor(s) for added users. This is a comma-separated list of values, e.g. mySuper.one,mySuper.two [default: None]

-U UDF_NAME, --udf UDF_NAME
If not specified in the defaults file, use this for the User Defined Field. [default: None]

-u USER, --user USER  If not specified in the defaults file, use -u to specify the xmatters user id that has permissions to get/update Site, Users, Devices, and Groups.

-V, --version         show program's version number and exit

-v                    set verbosity level. Each occurrence of v increases the logging level. By default it is ERRORs only, a single v (-v) means add WARNING logging, a double v (-vv) means add INFO logging, and a tripple v (-vvv) means add DEBUG logging [default: 0]

-x XMOD_URL, --xmodurl XMOD_URL
If not specified in the defaults file, use -i to specify the base URL of your xmatters instance. For example, 'https://myco.hosted.xmatters.com' without quotes.

```
* You can add multiple "v"'s to the -v command line option.  
   * A single "-v" means only show errors and warnings
   * A double "-vv" means to show errors, warnings, and info statements
   * A tripple "-vvv" means to show errors, warnings, info, and debug statements
