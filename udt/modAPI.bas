Attribute VB_Name = "modAPI"
'-- Original API Declare:
'-- Declare Function SetLocalTime Lib "Coredll" (lpSystemTime As SYSTEMTIME) As Integer

'-- API Declare to Set Local Time
'-- Notice this API requires a UDT as seen above!
Declare Function SetLocalTime Lib "Coredll" (ByVal lpSystemTime As String) As Long

