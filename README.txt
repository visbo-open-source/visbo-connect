The program visbo-connect is designed to work as a connector between the
VISBO web site and the VISBO clients (in the first step VISBO Project Edit, in short SPE)

Inspired by https://github.com/oikonomopo/electron-deep-linking-mac-win/blob/master/main.js
This example used work also for Windows.
In Windows you need to build an installation exe and it needs to be installed once.

It is implememented as a electron client in JavaScript.

First the executable gets called once for installation. (visbo-connect Setup x.x.x.exe)
Afterwards it will catch all the URL calls starting with visbo-connect://
For VISBO Proejct Edit it will handle the URLs visbo-connect://edit?vpid=...&vpvid=...&ott=...
The parameters after the ? will be extracted and cleaned for the SPE in the following way:
/e/vpid=.../vpvid=.../ott=.../
Afterwards visbo-connect calls Microsoft Excel with the cleaned parameter and the SPE sheet as the last parameter.
This enables SPE to opens project version for editing.
The One Time Token (ott) is used to login to the rest server without asking her for credetionals. It will use
the user, which is already logged in a web session and calls the Project Edit Client.

Logging to be described

Configuration of SPE paths to be described
