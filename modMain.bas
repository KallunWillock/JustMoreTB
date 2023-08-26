Attribute VB_Name = "modMain"
Option Explicit

' This is a simple workbook designed to log what I've learnt about using TwinBasic to extend VBA functionality.
' The code below is a near-carbon copy of the Standard DLL demonstrate available with TwinBasic (beta)
    
' TWINBASIC:
'
'    Module MainModule
'
'        [ DllExport ]
'        Public Function MultiplyMe(ByVal Number1 As Long, ByVal Number2 As Long) As Long
'            Return Number1 * Number2
'        End Function
'
'    End Module


#If Win64 Then
    Private Declare PtrSafe Function MultiplyMe Lib "DemoDLL_win64.dll" (ByVal Number1 As Long, ByVal Number2 As Long) As Long
#Else
    Private Declare Function MultiplyMe Lib "DemoDLL_win32.dll" (ByVal Number1 As Long, ByVal Number2 As Long) As Long
#End If

Sub TestMultiplyMe()
    
    Dim FirstNumber As Long, SecondNumber As Long, Result As Long
    
    FirstNumber = 111
    SecondNumber = 222
    
    Result = MultiplyMe(FirstNumber, SecondNumber)
    
    Debug.Print Result

End Sub
