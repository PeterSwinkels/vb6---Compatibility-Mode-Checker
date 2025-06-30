Attribute VB_Name = "Module1"
'This module contains this program's core procedures.
Option Explicit

'The Microsoft Windows API constants and functions used by this program.
Private Declare Function RegCloseKey Lib "Advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyExA Lib "Advapi32.dll" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueExA Lib "Advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long

Private Const ERROR_SUCCESS As Long = 0
Private Const HKEY_CURRENT_USER As Long = &H80000001
Private Const KEY_READ As Long = &H20019
Private Const REG_SZ As Long = 1&

Private Const MAX_REG_VALUE_DATA As Long = &HFFFFF   'Defines the maximum length in bytes allowed for a long string.

'This procedure is executed when this program is started.
Public Sub Main()
Dim KeyH As Long
Dim Data As String
Dim Length As Long
Dim ProgramPath As String

   ProgramPath = App.Path
   If Not Right$(ProgramPath, 1) = "\" Then ProgramPath = ProgramPath & "\"
   ProgramPath = ProgramPath & App.EXEName & ".exe"
   
   ProgramPath = Trim$(InputBox$("Check the compatibility mode for: ", , ProgramPath))

   If Not ProgramPath = vbNullString Then
      If RegOpenKeyExA(HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers", 0, KEY_READ, KeyH) = ERROR_SUCCESS Then
         Data = String$(MAX_REG_VALUE_DATA, vbNullChar)
         Length = Len(Data)
   
         If RegQueryValueExA(KeyH, ProgramPath, CLng(0), REG_SZ, Data, Length) = ERROR_SUCCESS Then
            Data = Left$(Data, Length - 1)
            MsgBox "Compatibility mode: " & vbCr & """" & Trim(Data) & """", vbInformation
         Else
            MsgBox "No compatibility mode found for the specified program.", vbInformation
         End If
   
         RegCloseKey KeyH
      Else
         MsgBox "Cannot access the relevant part of the registry.", vbExclamation
      End If
   End If
End Sub


