'Create the UFT application object
Set uftApp = CreateObject("QuickTest.Application") 

'Launch UFT
uftApp.Launch 

'Maximize and Make UFT visible
uftApp.Visible = True 
uftApp.WindowState = "Maximized"

'Open the UFT test
uftApp.Open "C:\SFAutomation\Scripts\LOI_Submission_RA_LOI_And_Application_001_02", True  

'Set run settings for the test 
Set uftTest = uftApp.Test 
'Continue test even though error occurs 
uftTest.Settings.Run.OnError = "NextStep"

'Run the UFT test
uftTest.Run 

'Close the test and Quit UFT

uftTest.Close 
uftApp.quit 

'Release the resources 
Set uftTest = Nothing 
Set uftApp = Nothing 