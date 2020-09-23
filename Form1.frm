VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compile"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   4380
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option2 
      Caption         =   "Second Sample"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox T2 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   1935
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "Form1.frx":0000
      Top             =   2640
      Width           =   4095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "First Sample"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.TextBox T1 
      Appearance      =   0  'Flat
      Height          =   1935
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "Form1.frx":004A
      Top             =   360
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Compile"
      Height          =   255
      Left            =   3240
      TabIndex        =   0
      Top             =   4680
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub DoIt(sFil As String)
    If sFil = "" Then Exit Sub
Dim v_Assembler As New clsAssembly
Dim v_ASM As New clsASM

    With v_Assembler
        .FileName = sFil
        If .CreateFile = False Then GoTo ContinueAnyWay:

        With v_ASM
            .wJMP 27
            .wAddData "Hello World!"
            .wAddData "  Islam Adel!"
            
            .wMOV DX_m, 2, 1
            .wMOV AH_m, 9
            .wINT &H21
            .wMOV AH_m, 0
            .wINT &H16
            
            .wMOV AL_m, 13
            .wMOV AH_m, 14
            .wINT &H10
            
            .wMOV AL_m, 10
            .wMOV AH_m, 14
            .wINT &H10
            
            .wMOV DX_m, 15, 1
            .wMOV AH_m, 9
            .wINT &H21
            .wMOV AH_m, 0
            .wINT &H16
            
            .wRET
        End With
        
        .CloseFile
        
        If MsgBox("Run file?", vbQuestion + vbYesNo, "Run?") = vbYes Then Shell .FileName, vbNormalFocus
    End With

ContinueAnyWay:
Set v_Assembler = Nothing
Set v_ASM = Nothing
End Sub
Public Sub DoIt2(sFil As String)
    If sFil = "" Then Exit Sub
Dim v_Assembler As New clsAssembly
Dim v_ASM As New clsASM

    With v_Assembler
        .FileName = sFil
        If .CreateFile = False Then GoTo ContinueAnyWay:

        With v_ASM
            .wMOV AL_m, 9
            .wADD AL_a, 48
            .wMOV AH_m, 14
            .wINT &H10
            .wMOV AH_m, 0
            .wINT &H16
            
            .wRET
        End With
        
        .CloseFile
        
        If MsgBox("Run file?", vbQuestion + vbYesNo, "Run?") = vbYes Then Shell .FileName, vbNormalFocus
    End With

ContinueAnyWay:
Set v_Assembler = Nothing
Set v_ASM = Nothing
End Sub

Private Sub Command1_Click()
If Option1.Value = True Then
    DoIt Trim(InputBox("Enter file ame", "File Name", "C:\Documents and Settings\Secret_Search\Desktop\a.com"))
ElseIf Option2.Value = True Then
    DoIt2 Trim(InputBox("Enter file ame", "File Name", "C:\Documents and Settings\Secret_Search\Desktop\a.com"))
End If
End Sub


Private Sub Form_Load()

End Sub

Private Sub Option1_Click()
T1.Enabled = Not T1.Enabled
T2.Enabled = Not T1.Enabled
End Sub


Private Sub Option2_Click()
T1.Enabled = Not T1.Enabled
T2.Enabled = Not T1.Enabled
End Sub


