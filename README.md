# hello-world-vsto
Basic add-in for troubleshooting the scope of Microsoft Excel Visual Studio Tools for Office (VSTO) add-in issues

## Overview
This basic VSTO add-in is intended to help discern whether an issue affecting PI DataLink or PI Builder affects all VSTO add-ins and is thus outside the scope of the PI System. Logging is performed via [NLog](http://nlog-project.org/) for the following:
* [ThisAddIn_Startup](https://msdn.microsoft.com/en-us/library/bb157876.aspx)
* ThisAddIn_Shutdown
* loading the add-in's Ribbon resources
* initializing a custom task pane
* performing an action from a custom task pane
* performing an action from the Ribbon

By default, the **YYYY-MM-DD.log** is placed in **C:\Users\Public\logs** and will look similar to this:

![alt-text](https://github.com/osi-JeremyRoberts/hello-world-vsto/blob/master/readme/Log%20example.png)

However, the output location, debug level, and other settings are configurable via **NLog.config** as documented [here](https://github.com/nlog/NLog/wiki/Configuration-file).

## Installation
This add-in is deployed via [ClickOnce](https://msdn.microsoft.com/en-us/library/bb386179.aspx). To install it on your machine, copy
* Application Files folder
* setup.exe
* TestVSTOAddIn.vsto

to a local directory and launch setup.exe. The application is not signed, so you will receive a warning that the publisher cannot be verified:

![alt-text](https://github.com/osi-JeremyRoberts/hello-world-vsto/blob/master/readme/Publisher%20cannot%20be%20verified.png)

After clicking Install, you should see a message indicating success:

![alt-text](https://github.com/osi-JeremyRoberts/hello-world-vsto/blob/master/readme/Successful%20install.png)

At this point, the add-in should be available in Excel.

![alt-text](https://github.com/osi-JeremyRoberts/hello-world-vsto/blob/master/readme/AddIn%20screenshot.png)
