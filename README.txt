The program visbo-connect is designed to work as a connector between the
VISBO web site and the VISBO clients (in the first step VISBO Project Edit, in short SPE)

Inspired by https://github.com/oikonomopo/electron-deep-linking-mac-win/blob/master/main.js
This example used work also for Windows.
In Windows you need to build an installation exe and it needs to be installed once.

It is implemented as a electron client in JavaScript.

First the executable gets called once for installation. (visbo-connect Setup x.x.x.exe)
Afterwards it will catch all the URL calls starting with visbo-connect://
For VISBO Project Edit it will handle the URLs visbo-connect://edit?vpid=...&vpvid=...&ott=...
The parameters after the ? will be extracted and cleaned for the SPE in the following way:
/e/vpid=.../vpvid=.../ott=.../
Afterwards visbo-connect calls Microsoft Excel with the cleaned parameter and the SPE sheet as the last parameter.
This enables SPE to opens project version for editing.
The One Time Token (ott) is used to login to the rest server without asking her for credentials. It will use
the user account, which is already logged in a web session and calls the Project Edit Client.

Logging is done to the local folder %APPDATA%\VISBO\visbo-connect\logs (APPDATA = C:\Users\<username>\AppData\Roaming).
Every day a new log file is written, the old log file is backed up with a name added with a date (like visboconnect_2022-08-22.log).
3 backup log files will be kept, afterwards they will be deleted

VisboConnect needs to where the Microsoft Excel executable is located. It also needs to know, where VISBO Project Edit Excel sheet is located.
This information can be passed to VisboConnect with the help of a json configuration file name vcn_config.json.
It needs to be located under the folder %APPDATA%\Visbo\visbo-connect\
The json file has two parameters defined. excelExe and speSheet. The example content looks like this: (notice!: the "\" in json has to be "\\")
{"excelExe":"C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\EXCEL.EXE","speSheet":"C:\\Visbo\\VISBO SPE\\Visbo Project Edit.xlsx"}

Building, Test and Installation:

The solution is located in the VISBO GitHub under https://github.com/visbo-open-source/visbo-connect
It contains all the components for the creation of an setup.exe build for Visbo Connect.

After the cloning of the GitHub solution to your local folder the program can be started for testing with a call like this:
npm start "visbo-connect://edit?vpid:627a4a80c0bdb36bb7f65062&vpvid:627a5a1fc0bdb36bb7f65a22"
The call can be found in the exampleTestCall.bat
If the parameters are valid, this call should start the VISBO Project Edit

To create a VISBO Connect setup.exe for windows, the following command needs to be executed:
npm run dist
After the builing process is finished, the setup exceutable can be found in the subfolder dist.
It will be named visbo-connect Setup with a version number in its name like
visbo-connect Setup 0.8.2.exe
The execution of the setup program will install visbo-connect on a windows computer.
The installation does not need any administrative right for installation.
The name, version and icon of the executable can be configured in the file package.json

(To create a VISBO Connect MSI-Installer for windows look at the following link:
https://ourcodeworld.com/articles/read/927/how-to-create-a-msi-installer-in-windows-for-an-electron-framework-application)

After the sucessful installation the program can be called from the VISBO web site.

visbo-connect can be uninstalled with the normal windows dialogs.
