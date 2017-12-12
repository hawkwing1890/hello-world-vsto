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

By default, the **YYYY-MM-DD.log** is placed in **[basedir]\logs** and will look similar to this:

![alt-text](https://github.com/osi-JeremyRoberts/hello-world-vsto/blob/master/Log%20example.png)

However, the output location, debug level, and other settings are configurable via **NLog.config** as documented [here](https://github.com/nlog/NLog/wiki/Configuration-file).

## Prerequisites
