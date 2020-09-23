VERSION 5.00
Begin VB.Form frmFirstTime 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Niro Professional Address Book"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8265
   ControlBox      =   0   'False
   DrawMode        =   14  'Copy Pen
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   8265
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Next"
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
      Height          =   615
      Left            =   3045
      MouseIcon       =   "AddressBook2005frmFirstTime.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "This is the first time you run ""Niro Professional Address Book v 1.0"". Thank you for using our software."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   8025
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmFirstTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmPassword.Show
Unload Me
End Sub

Private Sub Form_Load()
App.TaskVisible = False
Dim db As Database
Dim rs As Recordset
Dim WS As Workspace

Set WS = DBEngine.Workspaces(0)
    DbFile = (App.Path & "\Database\AddressBook.mdb")
    PwdString = "swordfish"
Set db = DBEngine.OpenDatabase(DbFile, False, False, ";PWD=" & PwdString)
Set rs = db.OpenRecordset("Other", dbOpenTable)

If rs("FirstRun") = 1 Then
Unload Me
End If

If rs("PasswordProtected") = "Yes" Then
frmLogIn.Show
Exit Sub
ElseIf rs("PasswordProtected") = "No" Then

If rs("FirstRun") = 1 Then
frmMain.Show
Unload Me
End If

rs.Edit
rs("FirstRun") = 1
rs.Update
End If


End Sub
