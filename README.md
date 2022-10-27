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
.\venv\Scripts\pyinstaller --noconsole --add-data "static/*;static/" --distpath ..\ --name "IFRC Keyword Searcher" .\search_for_keywords.py
```
- ```.\venv\Scripts\pyinstaller``` path to pyinstaller in the virtual environment
- ```--noconsole``` does not open a console when the executable is run, only the application
- ```--add-data``` this is required to add the static files (logos)
- ```--distpath``` this specifies the location for the final executable file
- ```--name``` name to assign to the bundled app

You can add a ```--onefile``` which means everything is bundled into one executable file, however this slows down start-up of the GUI.

The application can be run by double-clicking the executable ```IFRC Keyword Searcher.exe``` file. With the above configuration it can take up to 30 seconds to load (described more [below](#loading-time-and-windows-security)). Once loaded it should appear like this:

<p align="center">
  <img src="static/app_screenshot.png?raw=true" alt="IFRC Keyword Searcher GUI application screenshot"/>
</p>

### Loading time and Windows security

The application takes some time to load. Without the ```--onefile``` argument, the application takes approximately 30s to load. With the ```--onefile``` argument it takes much longer (> 10m)! The cause of this delay is Windows malware scanning. If you want the application to load faster, you can turn this off by going to **Windows security &rarr; Virus and threat protection &rarr; Manage settings**, and clicking to turn off **Real time protection**. It is recommended that you turn this back on after the application has loaded.
