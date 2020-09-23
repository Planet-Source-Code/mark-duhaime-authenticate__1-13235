Attribute VB_Name = "Auth"
Option Explicit

'Vars for Authentication
Global AuthKey As Boolean
Global AuthString As String

' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

Const gREGKEYSYSINFOLOC = "System\CurrentControlSet\Control\ComputerName\ComputerName"
Const gREGKEYSYSINFO = "System\CurrentControlSet\Control\ComputerName\ComputerName"
Const gREGVALSYSINFO = "ComputerName"
Const RegKey = "Reg"
Global Register As String

'String to hold Registry Computer Name
Global SysInfoPath As String

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function GetVolumeInformation Lib "kernel32.dll" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Integer, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long

'Put your encryption string in the
'EncryptName Variable 20 characters
Global Const EncryptName = "Putencyptstringhere "

'Put your project name here
'This is an entry in the registry that is created
Const RegPath = "SOFTWARE\Project"

Public Sub StartSysInfo()
    
    On Error GoTo SysInfoErr
  
    ' Try To Get System Info Program Path\Name From Registry...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    '
    Else
        GoTo SysInfoErr
    End If
    
    Exit Sub

SysInfoErr:
    MsgBox "System Information Is Unavailable At This Time", vbOKOnly
    
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim I As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        KeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For I = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, I, 1)))   ' Build Value Char. By Char.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
    End Select
    
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:      ' Cleanup After An Error Has Occured...
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function

Sub InvertIt()
    Dim Temp As Integer
    Dim Hold As Integer
    Dim I As Integer
    Dim TempStr As String
        
    TempStr = ""
    For I = 1 To Len(SysInfoPath)
        Temp = Asc(Mid$(SysInfoPath, I, 1))
        Hold = 0
Top:
    Select Case Temp
        Case Is > 127
            Hold = Hold + 1
            Temp = Temp - 128
            GoTo Top
        Case Is > 63
            Hold = Hold + 2
            Temp = Temp - 64
            GoTo Top
        Case Is > 31
            Hold = Hold + 4
            Temp = Temp - 32
            GoTo Top
        Case Is > 15
            Hold = Hold + 8
            Temp = Temp - 16
            GoTo Top
        Case Is > 7
            Hold = Hold + 16
            Temp = Temp - 8
            GoTo Top
        Case Is > 3
            Hold = Hold + 32
            Temp = Temp - 4
            GoTo Top
        Case Is > 1
            Hold = Hold + 64
            Temp = Temp - 2
            GoTo Top
        Case Is = 1
            Hold = Hold + 128
            
    End Select
        Temp = 255 Xor Hold
        TempStr = TempStr + Chr(Temp)
    Next I
    
    SysInfoPath = TempStr
End Sub

Sub EncryptIt()
    Dim Temp As Integer
    Dim Temp1 As Integer
    Dim Hold As Integer
    Dim I As Integer
    Dim J As Integer
    Dim TempStr As String

    TempStr = ""
    For I = 1 To Len(EncryptName)
        Hold = 0
        Temp = Asc(Mid$(EncryptName, I, 1))
        For J = 1 To Len(SysInfoPath)
            Temp1 = Asc(Mid$(SysInfoPath, J, 1))
            Hold = Temp Xor Temp1
         Next J
        TempStr = TempStr + Chr(Hold)
    Next I
    
    SysInfoPath = TempStr
End Sub

Sub EncipherIt()
    Dim Temp As Integer
    Dim Hold As String
    Dim I As Integer
    Dim J As Integer
    Dim TempStr As String
    Dim Temp1 As String
    
    TempStr = ""
    For I = 1 To Len(SysInfoPath)
        Temp = Asc(Mid$(SysInfoPath, I, 1))
        Temp1 = Hex(Temp)
        If Len(Temp1) = 1 Then
            Temp1 = "0" & Temp1
        End If
        For J = 1 To 2
            Hold = Mid$(Temp1, J, 1)
            Select Case Hold
                Case "0"
                    TempStr = TempStr + "7"
                Case "1"
                    TempStr = TempStr + "B"
                Case "2"
                    TempStr = TempStr + "F"
                Case "3"
                    TempStr = TempStr + "D"
                Case "4"
                    TempStr = TempStr + "1"
                Case "5"
                    TempStr = TempStr + "9"
                Case "6"
                    TempStr = TempStr + "3"
                Case "7"
                    TempStr = TempStr + "A"
                Case "8"
                    TempStr = TempStr + "6"
                Case "9"
                    TempStr = TempStr + "5"
                Case "A"
                    TempStr = TempStr + "E"
                Case "B"
                    TempStr = TempStr + "8"
                Case "C"
                    TempStr = TempStr + "0"
                Case "D"
                    TempStr = TempStr + "C"
                Case "E"
                    TempStr = TempStr + "2"
                Case "F"
                    TempStr = TempStr + "4"
            End Select
        Next J
    Next I
    SysInfoPath = TempStr
End Sub

Public Sub GetSubKey()

    If Not GetKeyValue(HKEY_LOCAL_MACHINE, RegPath, RegKey, Register) Then
        'Rem Not in registry
        
    End If
    
End Sub

'GetSerialNumber Procedure - Put this in the module or form where it is called.
Function GetSerialNumber(strDrive As String) As Long
    Dim SerialNum As Long
    Dim Res As Long
    Dim Temp1 As String
    Dim Temp2 As String
    'initialise the strings
    Temp1 = String$(255, Chr$(0))
    Temp2 = String$(255, Chr$(0))
    'call the API function
    Res = GetVolumeInformation(strDrive, Temp1, Len(Temp1), SerialNum, 0, 0, Temp2, Len(Temp2))
    GetSerialNumber = SerialNum
    
End Function

Private Sub Main()
    'Dimension our variables
    Dim TempStr As String
    Dim RegStr As String
    Dim I As Integer
    Dim SerialNumber As Long
    
    'Get The Computer Name in the registry
    'StartSysInfo
    SerialNumber = GetSerialNumber("C:\")
    SysInfoPath = Str(SerialNumber)
    
    'For encrypting purposes make the length
    'of it no more than 20 character
    If Len(SysInfoPath) > 20 Then
        SysInfoPath = Left$(SysInfoPath, 20)
    End If
    'invert the computer name
    InvertIt
    EncryptIt
    EncipherIt
    GetSubKey
    
    'verify it
    If Len(SysInfoPath) > 20 Then
        SysInfoPath = Left$(SysInfoPath, 20)
    End If
    For I = 1 To Len(SysInfoPath)
        TempStr = Mid$(SysInfoPath, I, 1)
        Select Case TempStr
            Case "0"
                AuthString = AuthString + "G"
            Case "1"
                AuthString = AuthString + "I"
            Case "2"
                AuthString = AuthString + "K"
            Case "3"
                AuthString = AuthString + "M"
            Case "4"
                AuthString = AuthString + "O"
            Case "5"
                AuthString = AuthString + "Q"
            Case "6"
                AuthString = AuthString + "S"
            Case "7"
                AuthString = AuthString + "U"
            Case "8"
                AuthString = AuthString + "V"
            Case "9"
                AuthString = AuthString + "T"
            Case "A"
                AuthString = AuthString + "R"
            Case "B"
                AuthString = AuthString + "P"
            Case "C"
                AuthString = AuthString + "N"
            Case "D"
                AuthString = AuthString + "L"
            Case "E"
                AuthString = AuthString + "J"
            Case "F"
                AuthString = AuthString + "H"
        End Select
    Next I
    
    'Retrieve Registry info for verification
    Register = GetSetting("Project", "Options", "Auth")
    If Len(Register) = 0 Then
        'rem load form to enter #
        TempStr = Left$(SysInfoPath, 5)
        authen.txtAuth1.Text = TempStr
        TempStr = Mid$(SysInfoPath, 6, 5)
        authen.txtAuth2.Text = TempStr
        TempStr = Mid$(SysInfoPath, 11, 5)
        authen.txtAuth3.Text = TempStr
        TempStr = Right$(SysInfoPath, 5)
        authen.txtAuth4.Text = TempStr
        authen.Show 1
    Else
        If Register <> AuthString Then
            End
        End If
    End If
    
End Sub

