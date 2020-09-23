Attribute VB_Name = "EasyExt"
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hkey As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Public Function WriteExt(strPath As String, strValue As String, strdata As String)
    Dim hkey As Long
    hkey = &H80000000
    Dim y As Long
    Dim x As Long
    x = RegCreateKey(hkey, strPath, y)
    x = RegSetValueEx(y, strValue, 0, 1, ByVal strdata, Len(strdata))
    x = RegCloseKey(y)
    savestring = x
End Function

Public Sub AddExt(Ext As String, ExtIcon As String, ExtPath As String)
    Ext = Replace(Ext, ".", "")
    Call WriteExt("." & Ext, "", Ext & "file")
    Call WriteExt("." & Ext, "Content Type", "text/plain")
    Call WriteExt(Ext & "file", "", "tst")
    Call WriteExt(Ext & "file\DefaultIcon", "", ExtIcon)
    Call WriteExt(Ext & "file\Shell\Open", "", "")
    Call WriteExt(Ext & "file\Shell\Open\command", "", """" & ExtPath & """ /Open ""%1""")
End Sub

Public Sub AddMenu(Ext As String, ExtPath As String, ExtText As String, ExtCall As String)
    Ext = Replace(Ext, ".", "")
    Call WriteExt(Ext & "file\Shell\" & ExtText, "", ExtText)
    Call WriteExt(Ext & "file\Shell\" & ExtText & "\command", "", """" & ExtPath & """ " & ExtCall & " ""%1""")
End Sub
