# Outlook-Macros
My Outlook Macros

Option Explicit

Public Const VK_NUMLOCK = &H90

Public Declare Function GetKeyState Lib "user32" (ByVal VirtKey As Long) As Long

Sub NumLockState()
    MsgBox "Numlock is " & IIf(GetKeyState(VK_NUMLOCK) = 1, "On", "Off")
End Sub

Sub CategorySearch()
    'MsgBox "Test"
    
    SendKeys "^{e}"
    SendKeys "Categories: Mike"
    SendKeys "{ENTER}"
    
    'num lock fix
    If GetKeyState(VK_NUMLOCK) <> 1 Then
        SendKeys "{NUMLOCK}", True
    End If
    
End Sub

Sub DoneReply()
    Dim NumLockState As Long
    NumLockState = GetKeyState(VK_NUMLOCK)
    
    SendKeys "Done"
    
    'num lock fix
    If NumLockState <> GetKeyState(VK_NUMLOCK) Then
        SendKeys "{NUMLOCK}", True
    End If
    
End Sub

Sub EA_Rights()
    Dim NumLockState As Long
    NumLockState = GetKeyState(VK_NUMLOCK)
    
    SendKeys "You now have this right.  Log out of eAutomate then log back in for these changes to take effect."
    
    'num lock fix
    If NumLockState <> GetKeyState(VK_NUMLOCK) Then
        SendKeys "{NUMLOCK}", True
    End If
    
End Sub

Sub NetworkPassReset()
    Dim NumLockState As Long
    NumLockState = GetKeyState(VK_NUMLOCK)
    
    SendKeys "Network password reset to "
    SendKeys "^{b}"
    SendKeys " Pa$$1234 "
    SendKeys "^{b}"
    SendKeys "{ENTER}"
    SendKeys "You will be asked to change your password at the next login."
    
    'num lock fix
    If NumLockState <> GetKeyState(VK_NUMLOCK) Then
        SendKeys "{NUMLOCK}", True
    End If
    
End Sub

Sub TeleVantageUnlock()
    Dim NumLockState As Long
    NumLockState = GetKeyState(VK_NUMLOCK)
    
    SendKeys "Unlocked, temp pin is "
    SendKeys "^{b}"
    SendKeys "112233"
    SendKeys "^{b}"
    SendKeys "{ENTER}"
    SendKeys "You will be ask to change you pin at next login.  Must be numerical and six digits long."
    
    'num lock fix
    If NumLockState <> GetKeyState(VK_NUMLOCK) Then
        SendKeys "{NUMLOCK}", True
    End If
    
End Sub
