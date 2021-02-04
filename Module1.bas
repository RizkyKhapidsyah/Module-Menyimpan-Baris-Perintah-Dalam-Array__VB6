Attribute VB_Name = "Module1"
Option Explicit

Public Function RtnGetCommandLine(ArgArray() As String) _
   As Integer

'Purpose: Save all command line arguments to an array and return
'         Number of command line arguments.  Assumes command line
'         arguments are delimited by one space or a tab
'
'Parameters:  ArgArray: Uninitialized String Array in which
'             Command Line Arguments will be saved
'
'Returns:     Number of Command Line Arguments
'
'Example:
'Dim sArray() As String
'Dim iCtr As Integer

'MsgBox RtnGetCommandLine(sArray) 'displays number of command
'                                 line arguments
'For iCtr = 0 To UBound(sArray)
 '   Debug.Print sArray(iCtr) 'Outputs each argument
'Next

'NOTE: As written, will not work if there are more than 10
'Command line arguments, but you can change this easily if
'necessary.
'************************************************************
   'Declare variables.
   Dim C As String
   Dim CmdLine As String
   Dim CmdLnLen As Integer
   Dim InArg As Boolean
   Dim I As Integer
   Dim NumArgs As Integer

   'Initialise variables
   ReDim ArgArray(10)
   NumArgs = 0
   InArg = False
   'Get command line arguments.
   CmdLine = Command()
   CmdLnLen = Len(CmdLine)
   'Go thru command line character at a time.
   For I = 1 To CmdLnLen
      C = Mid(CmdLine, I, 1)
      'Test for space or tab.
      If (C <> " " And C <> vbTab) Then
         'Neither space nor tab.
         'Test if already in argument.
         If Not InArg Then
         'New argument begins.
         'Test for too many arguments.
            If NumArgs = 10 Then Exit For
            NumArgs = NumArgs + 1
            InArg = True
         End If
         'Concatenate character to current argument.
         ArgArray(NumArgs) = ArgArray(NumArgs) & C
      Else
         'Found a space or tab.
         'Set InArg flag to False.
         InArg = False
      End If
   Next I
   'Resize array just enough to hold arguments.
   ReDim Preserve ArgArray(NumArgs)
   RtnGetCommandLine = NumArgs

End Function

