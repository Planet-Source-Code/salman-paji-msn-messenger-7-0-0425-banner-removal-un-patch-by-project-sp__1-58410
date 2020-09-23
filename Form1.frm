VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MSN 7.0.0425 Banner Removal UN/Patch"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4455
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   4455
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "UnPatch"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   960
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Patch"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Project SP : http://www.ptdnet.com/"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   4215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "MSN Messenger 7.0.0425 Banner Removal UN/Patch by Project SP"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BR_Patch()
Dim sp As Integer
sp = FreeFile
Open "msnmsgr.exe" For Binary As #sp
Put #sp, &H5624F5, "n"
Put #sp, &H5624F7, "n"
Put #sp, &H5624F8, "e"
Put #sp, &H5624F9, " "
Put #sp, &H5624FA, " "
Put #sp, &H5624FB, " "
Put #sp, &H5624FC, " "
Put #sp, &H5624FD, " "
Put #sp, &H5624FE, " "
Put #sp, &H5624FF, " "
Put #sp, &H562500, " "
Put #sp, &H562501, " "
Put #sp, &H562502, " "
Put #sp, &H562503, " "
Put #sp, &H562504, " "
Put #sp, &H562505, " "
Put #sp, &H562506, " "
Put #sp, &H562507, " "
Put #sp, &H562508, " "
Put #sp, &H562509, " "
Put #sp, &H56250A, " "
Put #sp, &H56250B, " "
Put #sp, &H56250C, " "
Put #sp, &H56250D, " "
Put #sp, &H56250E, " "
Put #sp, &H56250F, " "
Put #sp, &H562510, " "
Put #sp, &H562511, " "
Put #sp, &H562512, " "
Put #sp, &H562513, " "
Put #sp, &H562514, " "
Put #sp, &H562515, " "
Put #sp, &H562516, " "
Put #sp, &H562517, " "
Put #sp, &H562518, " "
Put #sp, &H562519, " "
Put #sp, &H56251A, " "
Close #sp
End Sub
Private Sub BR_UnPatch()
Dim sp As Integer
sp = FreeFile
Open "msnmsgr.exe" For Binary As #sp
Put #sp, &H5624F5, "t"
Put #sp, &H5624F7, "p"
Put #sp, &H5624F8, " "
Put #sp, &H5624F9, "l"
Put #sp, &H5624FA, "a"
Put #sp, &H5624FB, "y"
Put #sp, &H5624FC, "o"
Put #sp, &H5624FD, "u"
Put #sp, &H5624FE, "t"
Put #sp, &H5624FF, "="
Put #sp, &H562500, "v"
Put #sp, &H562501, "e"
Put #sp, &H562502, "r"
Put #sp, &H562503, "t"
Put #sp, &H562504, "i"
Put #sp, &H562505, "c"
Put #sp, &H562506, "a"
Put #sp, &H562507, "l"
Put #sp, &H562508, "f"
Put #sp, &H562509, "l"
Put #sp, &H56250A, "o"
Put #sp, &H56250B, "w"
Put #sp, &H56250C, "l"
Put #sp, &H56250D, "a"
Put #sp, &H56250E, "y"
Put #sp, &H56250F, "o"
Put #sp, &H562510, "u"
Put #sp, &H562511, "t"
Put #sp, &H562512, "("
Put #sp, &H562513, "0"
Put #sp, &H562514, ","
Put #sp, &H562515, "2"
Put #sp, &H562516, ","
Put #sp, &H562517, "2"
Put #sp, &H562518, ","
Put #sp, &H562519, "2"
Put #sp, &H56251A, ")"
Close #sp
End Sub
Private Sub Command1_Click()
Call BR_Patch
End Sub
Private Sub Command2_Click()
Call BR_UnPatch
End Sub
