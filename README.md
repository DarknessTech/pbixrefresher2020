Power BI Refresher 2020
======
Script for automation of refreshing Power BI Desktop workbooks. Edited on Python 3.8 with pywinauto, pyautogui.

Added compatibility to PowerBI Version 2.84.981.0 (August 2020) on Windows 10 with English/Portuguese(Brazil) locales.

*This script originaly is not mine (original post bellow), I just updated it to make it run on the latest version of PowerBI.*

Based on (Scripts and Guides):
------
```
https://github.com/dubravcik/pbixrefresher-python (Original Script)
https://github.com/LevonPython/PbiRefresher
https://github.com/pywinauto/pywinauto/issues/943 
https://stackoverflow.com/questions/15887729/can-the-gui-of-an-rdp-session-remain-active-after-disconnect
```
Installation
------
Install Windows SDK (https://developer.microsoft.com/en-us/windows/downloads/sdk-archive/)

Install script using `pip`

```
pip install pbixrefresher
pip install pywinauto 
pip install pyautogui
```
Replace the pbixrefresher.py file in the python installation folder (C:\Users\USERNAME\AppData\Local\Programs\Python\Python38\Lib\site-packages\pbixrefresher) with the pbixrefresher.py file included in this release.

Usage
-----
Before runing the Script, run Inspect.exe (included in Windows SDK instalation folder), it makes possible for the script to detect the Power BI elements.

```
pbixrefresher <WORKBOOK> [-workspace <WORKSPACE>] [--refresh-timeout <REFRESH_TIMEOUT>] [--no-publish]

where <WORKBOOK> is path to .pbix file
      --workspace <text> is name of online Power BI service work space to publish in (default My workspace)
      --refresh-timeout <number> is time in seconds to wait to refresh end (default 30000)
      --no-publish is switch to just refresh and save the workbook and skip publishing to online service (default False)
      --init-wait <number> is time to wait until Power BI Desktop starts (default 60)
```

Reminder
-----
Please keep in mind that this script uses GUI of Power BI Desktop and it needs that a user is logged in Windows session. You should also deactivate lock screen time. Ideally you should schedule the script on a computer where the GUI is not used to not interfere the scripting, for example dedicated Virtual Machine.

Running in a VM
-----

You can use the closesession.bat, just edit the username with yours. Instead of disconnecting from Remote Desktop normally, with closesession.bat it closes the session but maintein the GUI active so the python script can work.
