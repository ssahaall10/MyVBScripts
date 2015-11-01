'******************************************************************************
'Program Name : Prevent_AutoLock_v2.vbs
'Description  : This vbscript will prevent local system auto locking
'Created by   : Saurav Saha
'Date	      : 08/12/2014
'Notes	      : Please use in local system only. Not to be used in VMs.
'******************************************************************************
Option Explicit

call Prevent_AutoLock()

Function Prevent_AutoLock()
	'Variable declaration
	Dim wsc, myValue, i, flag, input, a
	input = null
	set wsc = CreateObject("WScript.Shell")
	
	input = InputBox("Enter number of hours","Enter Numbers Only" )
	myValue = 60*input
	'msgbox	 DateAdd("n",myValue,date())
	msgbox	"Your system will not get auto lock until " &FormatDateTime(DateAdd("n",myValue,now()),3)	&" of date "&FormatDateTime(DateAdd("n",myValue,now()),2)
	for i = 1 to myValue Step 1
		WScript.Sleep (60*1000)
		wsc.SendKeys ("{SCROLLLOCK 2}")
		'msgbox i+1 &"min"
		if i = myValue then
			a = msgbox ("Auto lock feature is now enabled. Your system will now get auto lock",1, "Want to re-run the script")
			if a = 1 then 
				call Prevent_AutoLock()
			else	
				WScript.Quit()
			end if	
		end if		
	next
		
End Function	