Attribute VB_Name = "INIReadWrite"
Option Explicit
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Dim m_File As String, m_Buffer As Long

Public Sub INISetup(FileName As String, BufferSize As Long)
   m_Buffer = BufferSize
   m_File = FileName
End Sub

Public Function Read_Ini(iSection As String, iKeyName As String, Optional iDefault As String)
    m_Buffer = 5000
    m_File = App.Path & "\config.ini"
    
    If m_Buffer = 0 Then
        Err.Raise "670", "INIReadWrite", "Buffer size not set"
    ElseIf m_File = "" Then
        Err.Raise "670", "INIReadWrite", "Filename not set"
    End If
    Dim Ret As String, NC As Long
    
    Ret = String(m_Buffer, 0)
    
    NC = GetPrivateProfileString(iSection, iKeyName, iDefault, Ret, m_Buffer, m_File)
    
    If NC <> 0 Then
        Ret = Left$(Ret, NC)
    Else
        Ret = ""
    End If
    
    Ret = Replace(Ret, "%%&&Chr(13)&&%%", vbCrLf)
    
    Read_Ini = Ret
End Function

Public Sub Write_Ini(iSection As String, iKeyName As String, iValue As Variant)
    If m_Buffer = 0 Then
        Err.Raise "670", "INIReadWrite", "Buffer size not set"
    ElseIf m_File = "" Then
        Err.Raise "670", "INIReadWrite", "Filename not set"
    End If
    
    iValue = CStr(iValue)
    
    iValue = Replace(iValue, vbCrLf, "%%&&Chr(13)&&%%")
    WritePrivateProfileString iSection, iKeyName, CStr(iValue), m_File
End Sub

Public Function Read_Sections()
    If m_Buffer = 0 Then
        Err.Raise "670", "INIReadWrite", "Buffer size not set"
    ElseIf m_File = "" Then
        Err.Raise "670", "INIReadWrite", "Filename not set"
    End If
    Dim Ret As String, NC As Long
    
    Ret = String(m_Buffer, 0)
    
    NC = GetPrivateProfileString(vbNullString, vbNullString, vbNullString, Ret, m_Buffer, m_File)
    
    If NC <> 0 Then
        Ret = Left$(Ret, NC - 1)
    End If
    
    Read_Sections = Ret
End Function

Public Function Read_Keys(iSection As String)
    If m_Buffer = 0 Then
        Err.Raise "670", "INIReadWrite", "Buffer size not set"
    ElseIf m_File = "" Then
        Err.Raise "670", "INIReadWrite", "Filename not set"
    End If
    Dim Ret As String, NC As Long
    
    Ret = String(m_Buffer, 0)
    
    NC = GetPrivateProfileString(iSection, vbNullString, vbNullString, Ret, m_Buffer, m_File)
    
    If NC <> 0 Then
        Ret = Left$(Ret, NC - 1)
    End If
    Read_Keys = Ret
End Function

Public Function DeleteSection(iSection As String)
    If m_Buffer = 0 Then
        Err.Raise "670", "INIReadWrite", "Buffer size not set"
    ElseIf m_File = "" Then
        Err.Raise "670", "INIReadWrite", "Filename not set"
    End If
    WritePrivateProfileString iSection, vbNullString, vbNullString, m_File
End Function

Function DeleteKey(iSection As String, iKeyName As String)
    If m_Buffer = 0 Then
        Err.Raise "670", "INIReadWrite", "Buffer size not set"
    ElseIf m_File = "" Then
        Err.Raise "670", "INIReadWrite", "Filename not set"
    End If
    WritePrivateProfileString iSection, iKeyName, vbNullString, m_File
End Function

