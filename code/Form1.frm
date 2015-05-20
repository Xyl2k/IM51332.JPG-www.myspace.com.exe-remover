VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Désinfecteur: IM51332.JPG-www.myspace.com.exe"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8535
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   8535
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   2655
      Left            =   120
      TabIndex        =   17
      Top             =   2040
      Visible         =   0   'False
      Width           =   8295
      Begin VB.CommandButton Command9 
         Caption         =   "Modifier"
         Height          =   495
         Left            =   240
         TabIndex        =   23
         Top             =   2040
         Width           =   6855
      End
      Begin VB.CommandButton Command8 
         Caption         =   "[ OK ]"
         Height          =   495
         Left            =   7200
         TabIndex        =   22
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   240
         TabIndex        =   21
         Text            =   "http://www.google.com/webhp?hl=xx-hacker"
         Top             =   1320
         Width           =   7815
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   600
         Width           =   7815
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Nouvelle valeur:"
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Valeur actuel:"
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton Command7 
      Caption         =   "IE Start page"
      Enabled         =   0   'False
      Height          =   615
      Left            =   120
      TabIndex        =   16
      Top             =   840
      Width           =   2175
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Détruire key3"
      Height          =   615
      Left            =   6480
      TabIndex        =   15
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Détruire key2"
      Height          =   615
      Left            =   4440
      TabIndex        =   14
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Détruire IM51332.JPG-www.myspace.com.exe"
      Height          =   615
      Left            =   4320
      TabIndex        =   7
      Top             =   120
      Width           =   4095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "[ Rafraîchissement ]"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   8295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Détruire infocard.exe"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Détruire key1"
      Height          =   615
      Left            =   2400
      TabIndex        =   1
      Top             =   840
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   8295
      Begin VB.Label Label11 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   2280
         Width           =   8055
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   585
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   585
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   585
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   585
      End
      Begin VB.Label Label8 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   8055
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   4440
      TabIndex        =   8
      Top             =   7080
      Width           =   975
   End
   Begin VB.Label userweak 
      Caption         =   "Label2"
      Height          =   255
      Left            =   3480
      TabIndex        =   5
      Top             =   6720
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Label2"
      Height          =   255
      Left            =   3000
      TabIndex        =   4
      Top             =   6360
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Label1"
      Height          =   255
      Left            =   2880
      TabIndex        =   3
      Top             =   6000
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
     Dim c As New cRegistry
 
     






Private Sub Command1_Click()
 Dim r
 Set r = CreateObject("WScript.Shell")
 r.RegDelete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\Firewall Administrating"
Command1.Enabled = False
Call Form_Load
End Sub

Private Sub Command4_Click()

Kill ("C:/Documents and Settings/" & userweak.Caption & "/Bureau/IM51332.JPG-www.myspace.com.exe")
Command4.Enabled = False
Call Form_Load
End Sub







Private Sub Command5_Click()
Call Form_Load
End Sub

Private Sub Command6_Click()

Dim r
 Set r = CreateObject("WScript.Shell")
 r.RegDelete "HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Services\SharedAccess\Parameters\FirewallPolicy\StandardProfile\AuthorizedApplications\List\"

With c
'START EDIT CODE
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Services\SharedAccess\Parameters\FirewallPolicy\StandardProfile\AuthorizedApplications\List" 'le lien vers la clé voulue
        .ValueType = REG_SZ 'ou REG_DWORD s'il s'agit de chiffres seulement
        .ValueKey = "C:\Documents and Settings\" & userweak.Caption & "\Bureau\IM51332.JPG-www.myspace.com.exe" 'nom d'un key (lol)
        .Value = "" 'attribuer la valeur du clé au texte (affichage)
'STOP EDIT CODE
         End With
Command6.Enabled = False
Call Form_Load
End Sub

Private Sub Command7_Click()
Frame2.Visible = True

         
End Sub

Private Sub Command8_Click()
Frame2.Visible = False
Call Form_Load
End Sub

Private Sub Command9_Click()
With c
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = "Software\Microsoft\Internet Explorer\Main" 'le lien vers la clé voulue
        .ValueType = REG_SZ 'ou REG_DWORD s'il s'agit de chiffres seulement
        .ValueKey = "Start Page" 'nom d'un key (lol)
        .Value = Text2.Text 'the above key value(serialNum)
End With
With c
'START EDIT CODE
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = "Software\Microsoft\Internet Explorer\Main" 'le lien vers la clé voulue
        .ValueType = REG_SZ 'ou REG_DWORD s'il s'agit de chiffres seulement
        .ValueKey = "Start Page" 'nom d'un key (lol)
         Text1.Text = .Value 'attribuer la valeur du clé au texte (affichage)
'STOP EDIT CODE
         End With
         
End Sub

Private Sub Form_Load()
KillProcessus "infocard.exe"
Get_User_Name
FileExists ("C:/Documents and Settings/" & userweak.Caption & "/Bureau/IM51332.JPG-www.myspace.com.exe")
FileExists2 ("c:/windows/infocard.exe")
Call BooZa
End Sub

Private Sub Timer1_Timer()
    Const OriginalCaption = "Désinfecteur: IM51332.JPG-www.myspace.com.exe"
    Const ScrolledCaption = OriginalCaption & " From Xylibox - xylibox.free.fr "
    Static Position As Long
    If Position Then
        If Position >= Len(ScrolledCaption) Then
            Position = 0
            Me.Caption = OriginalCaption
            Timer1.Interval = 10000
        Else
            Position = Position + 1
            Me.Caption = Left(Right(ScrolledCaption, Len(ScrolledCaption) - Position) & Left(ScrolledCaption, Position), Len(OriginalCaption))
            Timer1.Interval = 100
        End If
    Else
        Position = Position + 1
        Timer1.Interval = 100
    End If
End Sub

Private Function FileExists(FullFileName As String) As Boolean

    On Error GoTo MakeF
        'If file does Not exist, there will be an Error
        Open FullFileName For Input As #1
        Close #1
        'no error, file exists
        Command4.Enabled = True
Label2.Caption = "[IM51332.JPG-www.myspace.com.exe] -> Found !"
Label2.ForeColor = &HFF&
SetAttr "C:/Documents and Settings/" & userweak.Caption & "/Bureau/IM51332.JPG-www.myspace.com.exe", vbNormal 'attrib -s -h -r filename.ext
         
         Exit Function
        FileExists = True
   
    Exit Function
MakeF:
        

Label2.Caption = "[IM51332.JPG-www.myspace.com.exe] -> Not Found"
Label2.ForeColor = &HC000&
Command4.Enabled = False
        FileExists = False
 

    Exit Function
End Function

Private Function FileExists2(FullFileName As String) As Boolean

    On Error GoTo MakeF
        'If file does Not exist, there will be an Error
        Open FullFileName For Input As #1
        Close #1
        'no error, file exists
Label3.Caption = "[infocard.exe] -> Found !"
Label3.ForeColor = &HFF&
Command3.Enabled = True
SetAttr "c:/windows/infocard.exe", vbNormal 'attrib -s -h -r filename.ext
        FileExists2 = True
    Exit Function
MakeF:
        'error, file does Not exist
        FileExists2 = False
        Command3.Enabled = False

       Label3.Caption = "[infocard.exe] -> Not Found"
Label3.ForeColor = &HC000&
Command3.Enabled = False

    Exit Function
End Function

Private Sub Command2_Click()
 Dim r
 Set r = CreateObject("WScript.Shell")
 r.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run\Firewall Administrating"
Command2.Enabled = False
Call Form_Load
End Sub

Private Sub Command3_Click()
Kill ("c:/windows/infocard.exe")
Command3.Enabled = False
Call Form_Load
End Sub

Private Function BooZa()
With c
'START EDIT CODE
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Run" 'le lien vers la clé voulue
        .ValueType = REG_SZ 'ou REG_DWORD s'il s'agit de chiffres seulement
        .ValueKey = "Firewall Administrating" 'nom d'un key (lol)
         Label4.Caption = .Value 'attribuer la valeur du clé au texte (affichage)
'STOP EDIT CODE
         End With
         
         With c
'START EDIT CODE
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = "Software\Microsoft\Windows\CurrentVersion\Run" 'le lien vers la clé voulue
        .ValueType = REG_SZ 'ou REG_DWORD s'il s'agit de chiffres seulement
        .ValueKey = "Firewall Administrating" 'nom d'un key (lol)
         Label5.Caption = .Value 'attribuer la valeur du clé au texte (affichage)
'STOP EDIT CODE
         End With
         
         With c
'START EDIT CODE
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "SYSTEM\ControlSet001\Services\SharedAccess\Parameters\FirewallPolicy\StandardProfile\AuthorizedApplications\List" 'le lien vers la clé voulue
        .ValueType = REG_SZ 'ou REG_DWORD s'il s'agit de chiffres seulement
        .ValueKey = "C:\Documents and Settings\" & userweak.Caption & "\Bureau\IM51332.JPG-www.myspace.com.exe" 'nom d'un key (lol)
         Label1.Caption = .Value 'attribuer la valeur du clé au texte (affichage)
'STOP EDIT CODE
         End With
         
         With c
'START EDIT CODE
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = "Software\Microsoft\Internet Explorer\Main" 'le lien vers la clé voulue
        .ValueType = REG_SZ 'ou REG_DWORD s'il s'agit de chiffres seulement
        .ValueKey = "Start Page" 'nom d'un key (lol)
         Text1.Text = .Value 'attribuer la valeur du clé au texte (affichage)
'STOP EDIT CODE
         End With
         
         If Label4.Caption = "" Then
         Label6.Caption = "(Key2) [HKEY_LOCAL_MACHINE] Firewall Administrating -> Not Found"
Label6.ForeColor = &HC000&
Command1.Enabled = False

Else
Label6.Caption = "(Key2) [HKEY_LOCAL_MACHINE] Firewall Administrating -> Found !"
Label6.ForeColor = &HFF&
Command1.Enabled = True
 End If
          If Label5.Caption = "" Then
         Label7.Caption = "(Key1) [HKEY_CURRENT_USER] Firewall Administrating -> Not Found"
Label7.ForeColor = &HC000&
Command2.Enabled = False
Else
Label7.Caption = "(Key1) [HKEY_CURRENT_USER] Firewall Administrating -> Found !"
Label7.ForeColor = &HFF&
Command2.Enabled = True

 End If
 
 
           If Label1.Caption = "" Then
         Label8.Caption = "(Key3) [HKEY_LOCAL_MACHINE] C:\Documents and Settings\" & userweak.Caption & "\Bureau\IM51332.JPG-www.myspace.com.exe -> Not Found"
Label8.ForeColor = &HC000&
Command6.Enabled = False

Else
Label8.Caption = "(Key3) [HKEY_LOCAL_MACHINE] C:\Documents and Settings\" & userweak.Caption & "\Bureau\IM51332.JPG-www.myspace.com.exe -> Found !"
Label8.ForeColor = &HFF&
Command6.Enabled = True

 End If
If Text1.Text = "http://www.gllod.com" Then
Label11.Caption = "[Internet Explorer] Start page changed !"
Label11.ForeColor = &HFF&
Command7.Enabled = True
Else
Label11.Caption = "[Internet Explorer] Start page not changed"
Label11.ForeColor = &HC000&
Command7.Enabled = False

 End If
End Function


