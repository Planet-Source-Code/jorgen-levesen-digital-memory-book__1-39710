Attribute VB_Name = "InstaKid"
Option Explicit
Public Declare Function OpenProcess Lib "kernel32" _
    (ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, _
    ByVal dwProcessId As Long) As Long
Public Declare Function GetExitCodeProcess Lib "kernel32" _
    (ByVal hProcess As Long, lpExitCode As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" _
    (ByVal hObject As Long) As Long
Public Const PROCESS_QUERY_INFORMATION = &H400


