VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4725
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   4875
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   HasDC           =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   4725
   ScaleWidth      =   4875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   1560
      Top             =   360
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' shadow by EBArtSoft

Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion      As Long
    dwMinorVersion      As Long
    dwBuildNumber       As Long
    dwPlatformId        As Long
    szCSDVersion        As String * 128
End Type

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long


Private Const SND_APPLICATION = &H80
Private Const SND_ALIAS = &H10000
Private Const SND_ALIAS_ID = &H110000
Private Const SND_ASYNC = &H1
Private Const SND_FILENAME = &H20000
Private Const SND_LOOP = &H8
Private Const SND_MEMORY = &H4
Private Const SND_NODEFAULT = &H2
Private Const SND_NOSTOP = &H10
Private Const SND_NOWAIT = &H2000
Private Const SND_PURGE = &H40
Private Const SND_RESOURCE = &H40004
Private Const SND_SYNC = &H0

Private Declare Function PlaySound Lib "WINMM.DLL" Alias "PlaySoundA" (ByRef Sound As Any, ByVal hLib As Long, ByVal lngFlag As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Private Sub Form_Click()
    Unload Me
    Form1.Show
End Sub

Sub SetTransparency(ByVal SrcHwnd As Long, ByVal Alpha As Byte)
    Dim Os    As OSVERSIONINFO
    Dim Style As Long
    Dim Ver   As Long
    Os.dwOSVersionInfoSize = Len(Os)
    GetVersionEx Os
    If Os.dwMajorVersion >= 5 And Os.dwMinorVersion >= 1 Then
        If Alpha = 255 Then
            Style = GetWindowLong(SrcHwnd, GWL_EXSTYLE)
            Style = Style And Not WS_EX_LAYERED
            SetWindowLong SrcHwnd, GWL_EXSTYLE, Style
        Else
            Style = GetWindowLong(SrcHwnd, GWL_EXSTYLE)
            Style = Style Or WS_EX_LAYERED
            SetWindowLong SrcHwnd, GWL_EXSTYLE, Style
            SetLayeredWindowAttributes SrcHwnd, 0, Alpha, LWA_ALPHA
        End If
    End If
End Sub

Private Sub Form_Load()
    SetTransparency Me.hwnd, 0
End Sub

Sub PlayOutro()
   Dim bSound() As Byte
   bSound = LoadResData(101, "WAVE")
   PlaySound bSound(0), 0, SND_MEMORY Or SND_NODEFAULT Or SND_SYNC Or SND_NOWAIT
   Erase bSound
End Sub

Private Sub Timer1_Timer()
    Static Alpha As Integer
    Dim i As Integer
    Alpha = Alpha + 8
    If Alpha > 255 Then
        SetTransparency Me.hwnd, 255
        If Alpha > 512 Then PlayOutro: Unload Me: Form1.Show
    Else
        SetTransparency Me.hwnd, Alpha
    End If
End Sub


