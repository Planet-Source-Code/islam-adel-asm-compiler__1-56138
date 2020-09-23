Attribute VB_Name = "modASM"
Public mFileNumber As Integer

Public Enum vADDRegs
    AX_a = &H5
    AL_a = &H4
    AH_a = &H80
    
    BX_a = &H83
    BL_a = &H80
    BH_a = &H80
    
    CX_a = &H83
    CL_a = &H80
    CH_a = &H80
    
    DX_a = &H83
    DL_a = &H80
    DH_a = &H80
End Enum

Public Enum vCMPRegs
    AX_c = &H3D
    AL_c = &H3C
    AH_c = &H80
    
    BX_c = &H83
    BL_c = &H80
    BH_c = &H80
    
    CX_c = &H83
    CL_c = &H80
    CH_c = &H80
    
    DX_c = &H83
    DL_c = &H80
    DH_c = &H80
End Enum

Public Enum vMOVRegs
    AX_m = &HB8
    AL_m = &HB0
    AH_m = &HB4
    
    BX_m = &HBB
    BL_m = &HB3
    BH_m = &HB7
    
    CX_m = &HB9
    CL_m = &HB1
    CH_m = &HB5
    
    DX_m = &HBA
    DL_m = &HB2
    DH_m = &HB6
End Enum

Public Enum vINCRegs
    AX_i = &H40
    AL_i = &HFE
    AH_i = &HFE
    
    BX_i = &H43
    BL_i = &HFE
    BH_i = &HFE
    
    CX_i = &H41
    CL_i = &HFE
    CH_i = &HFE
    
    DX_i = &H42
    DL_i = &HFE
    DH_i = &HFE
End Enum

Public Enum vPUSHRegs
    AX_pu = &H50
    BX_pu = &H53
    CX_pu = &H51
    DX_pu = &H52
End Enum

Public Enum vPOPRegs
    AX_po = &H58
    BX_po = &H5B
    CX_po = &H59
    DX_po = &H5A
End Enum
Public Function HiByte(WordIn As Integer) As Byte
    If WordIn% And &H8000 Then
        HiByte = &H80 Or ((WordIn% And &H7FFF) \ &HFF)
    Else
        HiByte = WordIn% \ 256
    End If
End Function
Public Function LoByte(WordIn As Integer) As Byte
    LoByte = WordIn% And &HFF&
End Function

