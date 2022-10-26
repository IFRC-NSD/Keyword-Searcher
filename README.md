# Keyword Searcher GUI

GUI application to run keyword searching on IFRC documents.

## Requirements

The GUI application must be generated on Windows in order that the final executable file can be run on Windows. Therefore all instructions below are given for Windows OS.

## Setup

To use this package, you should first set up a virtual environment:

```bash
python3 -m venv venv
```

Once you've created the virtual environment, you need to activate it:

```console
venv\Scripts\activate
```

If you get an error similar to ```venv\Scripts\Activate.ps1 cannot be loaded because running scripts is disabled on this system```, run the following (as suggested by [Microsoft Tech Support](https://social.technet.microsoft.com/Forums/windowsserver/en-US/964636ad-347e-4b23-8f7a-f36a558115dd/error-quotfile-cannot-be-loaded-because-the-execution-of-scripts-is-disabled-on-this-systemquot)). When prompted, select ```Y``` for ```Yes```. This should allow running ```venv``` in the current PowerShell session:

```console
Set-ExecutionPolicy Unrestricted -Scope Process
```

Finally, install the requirements from the ```requirements.txt``` into the virtual environment:

```bash
python -m pip install -r requirements.txt
```

To verify that the installation went well, show the installed packages with ```python -m pip list``` and ensure that ```pandas```, ```PyMuPDF```, and ```PySimpleGUI``` are listed.

## Usage

### Running the script directly

For testing and debugging, the GUI application can be started by running the ```search_for_keywords.py``` script from a Windows command prompt:

```bash
python .\search_for_keywords.py
```

### Generating and running the GUI application

To generate the GUI application, [PyInstaller](https://pyinstaller.org/en/stable/index.html) can be used (note this must be run on Windows so that the final executable can be run on Windows):

```bash
.\venv\Scripts\pyinstaller --onefile --noconsole --add-data "static/*;static/" --distpath ..\ .\search_for_keywords.py
```
- ```.\venv\Scripts\pyinstaller``` path to pyinstaller in the virtual environment
- ```--onefile``` creates a one-file bundled executable as this is more convenient for sharing
- ```--noconsole``` does not open a console when the executable is run, only the application
- ```--add-data``` this is required to add the static files (logos)
- ```--distpath``` this specifies the location for the final executable file
