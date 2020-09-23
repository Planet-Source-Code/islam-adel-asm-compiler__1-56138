VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAssembly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mFileName As String
Private mFileExist As Integer
Public Property Get FileName() As String
    FileName = mFileName
End Property

Public Property Let FileName(ByVal sFileName As String)
    mFileName = sFileName
End Property

Public Function CreateFile() As Boolean

    If mFileName = "" Then GoTo Er_CreateFile:
    If mFileNumber <> 0 Then GoTo Er_CreateFile:
    
    If mFileExists = True Then
        If MsgBox("The file " & mFileName & " already exists" & vbCrLf & "Do you want to overwrite the file?", vbExclamation + vbYesNo, "Assembly") = vbNo Then
            GoTo Er_CreateFile:
        Else
            Kill mFileName
        End If
    End If
    
    mFileNumber = FreeFile
    Open mFileName For Binary Access Write As #mFileNumber
    
    CreateFile = True

    Exit Function
Er_CreateFile:
    
    CreateFile = False

End Function
Private Function mFileExists() As Boolean

On Error GoTo Er_mFileExists:

    mFileExist = FreeFile
    Open mFileName For Input As #mFileExist
    Close #mFileExist
    mFileExist = 0
    
    mFileExists = True

    Exit Function

Er_mFileExists:

    mFileExists = False
End Function


Public Sub CloseFile()
    mFileNumber = 0
    Reset
    Close
End Sub

