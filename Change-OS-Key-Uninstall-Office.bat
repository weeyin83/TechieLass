::
:: Change Windows OS License Key, activate and un-install Office
:: Tested with Windows 7 and Office 2010
::
:: This does error reporting via the Event Viewer, events are created if the OS key is not activiated and if Office is not uninstalled
::
:: Author: Sarah Lean 14/09/2015 
:: www.techielass.com
::

::Change Windows OS License Key and activate
cscript //B "%windir%\system32\slmgr.vbs" /ipk 00000-00000-00000-00000-00000

cscript //B "%windir%\system32\slmgr.vbs" /ato

::Create Uninstall Office File
(
echo ^<Configuration Product="ProPlus"^>
echo  ^<Display Level="none" CompletionNotice="no" SuppressModal="yes" AcceptEula="yes" /^>
echo  ^<Setting Id="SETUP_REBOOT" Value="NEVER" /^>
echo ^</Configuration^>
) > "C:\Program Files (x86)\Common Files\microsoft shared\OFFICE14\Office Setup Controller\uninstall.xml"

::Uninstall Office

"C:\Program Files (x86)\Common Files\microsoft shared\OFFICE14\Office Setup Controller\setup.exe" /uninstall PROPLUS /config uninstall.xml 



::Check if Windows OS is activated and Office uninstalled
@echo off
cscript /nologo c:\windows\system32\slmgr.vbs /xpr | findstr /i /c:" will expire "> NUL 2>&1
if [%errorlevel%]==[0] (GOTO EXPIRING) else (GOTO OFFICECHECK)


:EXPIRING
eventcreate /t information /id 607 /l system /d "Windows OS License is not activated.  Please refer to Knowledge Base article for more information."

:OFFICECHECK

if exist "C:\program files(x86)\microsoft office" (eventcreate /t information /id 607 /l system /d "Office is still installed. Please refer to Knowledge Base article for more information.") else (GOTO END)

:END
