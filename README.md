# BeefJerkySupreme
Python Service for monitoring Abbyy batches in exception state

##Requirements
- pywin32 - for whichever version of Python you are running (notnecessarily the OS version!)
- pyodbc
- smtplib
- tabulate
- Abbyy Flexicapture (tested with version 11)
 

## Installation
Change necessary values in ExceptionService.py (namely, database calls, service names and email settings)
Run from an elevanted command prompt:

    python ExceptionsService.py install
Then to start:

    sc start ABBYYExceptionCheck
(or whatever you called the service on line #22)

Output is logged to %TEMP%\ExcetionSvc.log which, as services typically run under LOCALSYSTEM will be %WINDIR%\Temp

## Notes
The service is set to run manually - open SCM to set as Automatic (or Automatic Delayed) and to change the user account it runs under, if you so wish
