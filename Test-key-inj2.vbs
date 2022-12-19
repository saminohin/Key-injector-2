' Author Name: Hacker01 

Const delay = 600

' Define the KeyInjector class 
class KeyInjector 
  Private keypress 
  
  Public Sub OpenPowershell()
    Set keypress = CreateObject("WScript.shell") ' Create the keypress object 
    keypress.Run "powershell.exe" ' Open PowerShell 
    WScript.Sleep delay ' Wait 600 ms 
  End Sub 
  
  'Define the inject method 
  Public Sub Inject(stringCommand, Getdate)
    keypress.SendKeys stringCommand & "{ENTER}" & GetDate & "{ENTER}" ' Inject the command 
    WScript.Sleep delay ' Wait 600 ms 
  End Sub 

End class 

' Create an instance of the keyInjector class 
set injector = New KeyInjector

' Calls the function OpenpOwershell which just opens up the powershell terminal 
injector.OpenPowershell 

'Inject hello, world command and get injects Get-Date command wich gets the current date and time 
injector.inject "Write-Output 'Hello, world'", "Get-Date"