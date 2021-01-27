# Problems checker
 ZBX problems checker

# Installation

**In order to use this tool, you need to install python3 and chrome browser and keep it updated**: https://www.python.org/downloads/ and choose newest version of the installers for your OS.

Chrome download link: https://www.google.com/chrome/

![alt text](https://i.imgur.com/06EspWQ.png)

(select "Add to path" checkbox during installation)

## After installing Chrome and Python

**For Windows** from project directory in terminal run:
(To do that go to the folder this project is downloaded to and **Shift + right click** then select **Open command windows here**, then just simply paste the text from above (it will install all necessarry modules for python so everything works).

If you want to use global environment:
```
pip install -r requirements.txt
```
If you want to use virtual environment, in project folder:

Create and activate venv:
```
python -m venv c:\path\to\myenv
c:\path\to\myenv\Scripts\activate.bat
```

Then install requirements
```
pip install -r requirements.txt
```


**For ubuntu**  you have to:

```
pip3 install -r requirements.txt
```

# Configuration
Open configs.py that was transferred to you and **provide credencials for zbx_username and zbx_password**.
Also you can check if sheetnames are correct(names of the sheets where pos/bos error data is stored in excel files), they are stored in the list zbx_sheets in the same configs.py file


# Usage
In project structure you can see folder "Files", inside you can see two other folders:
Report - it's only used to temporarily save the report file downloaded from Zabbix, it is saved there until next script run if you'd need to check it.
Countries - Here you have to put all your Excel files so script goes through them all


Then you should just run the Main.py file normally, but preferably through CMD so you can catch errors, to do that, CD into project folder and:
```
python Main.py
```

