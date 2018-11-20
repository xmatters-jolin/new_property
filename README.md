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
Details of the installation go here. 

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
Run one of these commands:
* `python3 new_property.py -v -c -d defaults.json sites`
   * Processes Sites only
* `python3 new_property.py -v -c -d defaults.json admins`
   * Processes Admin Users only
* `python3 new_property.py -v -c -d defaults.json groups`
   * Processes Security Groups only
* `python3 new_property.py -v -c -d defaults.json all`
   * Processes Sites, Admin Users, and Security Groups
   

# Troubleshooting
* You can add multiple "v"'s to the -v command line option.  
   * A single "-v" means only show errors and warnings
   * A double "-vv" means to show errors, warnings, and info statements
   * A tripple "-vvv" means to show errors, warnings, info, and debug statements
