VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsASM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub wINT(m_Interrupt As Long)
    P Chr$(&HCD)
    P Chr$(m_Interrupt)
End Sub

Public Sub wMOV(vRegister As vMOVRegs, ParamArray m_Values() As Variant)
Dim mCount As Long

    P Chr$(vRegister)
    
    For mCount = LBound(m_Values()) To UBound(m_Values())
        PM CLng(m_Values(mCount))
    Next mCount
End Sub

Private Sub P(m_String As String)
    Put mFileNumber, , m_String
End Sub


Public Sub wAddData(sText As String, Optional sTerminator As String = "$")
    If Right(sText, 1) <> sTerminator Then sText = sText & sTerminator
    P sText
End Sub

Public Sub wPUSH(vRegister As vPUSHRegs)
    P Chr$(vRegister)
End Sub
Public Sub wPOP(vRegister As vPOPRegs)
    P Chr$(vRegister)
End Sub

Public Sub wJMP(m_Location As Long)
    P Chr$(&HEB)
    P Chr$(m_Location)
End Sub

Public Sub wRET()
    P Chr$(&HC3)
End Sub

Public Sub wCMP(vRegister As vCMPRegs, m_Value As Long)

    P Chr$(vRegister)

    Select Case vRegister
        Case vCMPRegs.AH_c
            P Chr$(&HFC)
        Case vCMPRegs.BH_c
            P Chr$(&HFF)
        Case vCMPRegs.BL_c
            P Chr$(&HFB)
        Case vCMPRegs.BX_c
            P Chr$(&HFB)
        Case vCMPRegs.CH_c
            P Chr$(&HFD)
        Case vCMPRegs.CL_c
            P Chr$(&HF9)
        Case vCMPRegs.CX_c
            P Chr$(&HF9)
        Case vCMPRegs.DH_c
            P Chr$(&HFE)
        Case vCMPRegs.DL_c
            P Chr$(&HFA)
        Case vCMPRegs.DX_c
            P Chr$(&HFA)
    End Select
End Sub

Public Sub wINC(vRegister As vINCRegs)
    P Chr$(vRegister)
    
    Select Case vRegister
        Case vINCRegs.AL_i
            P Chr$(&HC0)
        Case vINCRegs.AH_i
            P Chr$(&HC4)
        Case vINCRegs.BL_i
            P Chr$(&HC3)
        Case vINCRegs.BH_i
            P Chr$(&HC7)
        Case vINCRegs.CL_i
            P Chr$(&HC1)
        Case vINCRegs.CH_i
            P Chr$(&HC5)
        Case vINCRegs.DL_i
            P Chr$(&HC2)
        Case vINCRegs.DH_i
            P Chr$(&HC6)
    End Select

End Sub

Public Sub wJE(m_Location As Long)
    P Chr$(&H74)
    P Chr$(m_Location)
End Sub

Public Sub wADD(vRegister As vADDRegs, m_Value As Long)
    P Chr$(vRegister)
    
    Select Case vRegister
        Case vADDRegs.AH_a
            P Chr$(&HC4)
        Case vADDRegs.BH_a
            P Chr$(&HC7)
        Case vADDRegs.BX_a
            P Chr$(&HC3)
        Case vADDRegs.BL_a
            P Chr$(&HC3)
        Case vADDRegs.CH_a
            P Chr$(&HC5)
        Case vADDRegs.CL_a
            P Chr$(&HC1)
        Case vADDRegs.CX_a
            P Chr$(&HC1)
        Case vADDRegs.DH_a
            P Chr$(&HC6)
        Case vADDRegs.DL_a
            P Chr$(&HC2)
        Case vADDRegs.DX_a
            P Chr$(&HC2)
    End Select
    
    PM m_Value
End Sub

Public Sub PM(m_Long As Long)
        If CLng(m_Long) > 255 Then
            P Chr$(LoByte(CInt(m_Long)))
            P Chr$(HiByte(CInt(m_Long)))
        Else
            P Chr$(m_Long)
        End If
End Sub
