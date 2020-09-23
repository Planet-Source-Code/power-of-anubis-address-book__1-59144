VERSION 5.00
Begin VB.Form frmOptions 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Options"
   ClientHeight    =   2580
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6075
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2580
   ScaleWidth      =   6075
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Program Protection"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   930
      MouseIcon       =   "AddressBook2005frmOptions.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   960
      Width           =   4215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ok"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      MouseIcon       =   "AddressBook2005frmOptions.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2347
      TabIndex        =   0
      Top             =   240
      Width           =   1380
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
frmPassword.Show
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

