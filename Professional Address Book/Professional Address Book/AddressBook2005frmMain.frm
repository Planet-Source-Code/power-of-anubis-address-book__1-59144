VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Personal Address Book"
   ClientHeight    =   8745
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   10050
   Icon            =   "AddressBook2005frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8745
   ScaleWidth      =   10050
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3840
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Website description"
      Top             =   6840
      Width           =   6015
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6690
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   1080
      Width           =   3375
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   9240
      Picture         =   "AddressBook2005frmMain.frx":08CA
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblTotalContacts 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "99999"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   300
      Left            =   2280
      TabIndex        =   29
      Top             =   8280
      Width           =   795
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Contacts :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   28
      Top             =   8280
      Width           =   1920
   End
   Begin VB.Label lblFax 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   4560
      TabIndex        =   27
      Top             =   4680
      Width           =   5400
   End
   Begin VB.Label lblCity 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   300
      Left            =   4560
      TabIndex        =   26
      Top             =   2880
      Width           =   5295
   End
   Begin VB.Label lblBirthday 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   300
      Left            =   5040
      TabIndex        =   25
      Top             =   1800
      Width           =   4755
   End
   Begin VB.Label lblWebSite 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   5160
      TabIndex        =   24
      Top             =   5400
      Width           =   4740
   End
   Begin VB.Label lblNationality 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   300
      Left            =   5400
      TabIndex        =   23
      Top             =   1440
      Width           =   4425
   End
   Begin VB.Label lblEMAIL 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   4920
      TabIndex        =   22
      Top             =   5040
      Width           =   4995
   End
   Begin VB.Label lblCompany 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   5280
      TabIndex        =   21
      Top             =   5760
      Width           =   4635
   End
   Begin VB.Label lblMobile 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   4920
      TabIndex        =   20
      Top             =   4320
      Width           =   5025
   End
   Begin VB.Label lblPhoneNumber 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   5880
      TabIndex        =   19
      Top             =   3960
      Width           =   3975
   End
   Begin VB.Label lblCountry 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   300
      Left            =   5040
      TabIndex        =   18
      Top             =   2520
      Width           =   4815
   End
   Begin VB.Label lblAddress 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   300
      Left            =   5160
      TabIndex        =   17
      Top             =   3240
      Width           =   4755
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   300
      Left            =   4800
      TabIndex        =   16
      Top             =   1080
      Width           =   5010
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Comment :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3840
      TabIndex        =   15
      Top             =   6480
      Width           =   1305
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fax :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3840
      TabIndex        =   14
      Top             =   4680
      Width           =   600
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "City :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3840
      TabIndex        =   13
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Birthday :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3840
      TabIndex        =   12
      Top             =   1800
      Width           =   1155
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Web Site :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3840
      TabIndex        =   11
      Top             =   5400
      Width           =   1260
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nationality :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3840
      TabIndex        =   10
      Top             =   1440
      Width           =   1425
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3840
      TabIndex        =   9
      Top             =   5040
      Width           =   915
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Company :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3840
      TabIndex        =   8
      Top             =   5760
      Width           =   1275
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3840
      TabIndex        =   7
      Top             =   4320
      Width           =   945
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone Number :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3840
      TabIndex        =   6
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Country :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3840
      TabIndex        =   5
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3840
      TabIndex        =   4
      Top             =   3240
      Width           =   1155
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3840
      TabIndex        =   3
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contacts :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   1395
   End
   Begin VB.Menu mnuContacts 
      Caption         =   "&Contacts"
      Begin VB.Menu mnuNewContact 
         Caption         =   "New Contact"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuEditSelectedContact 
         Caption         =   "Edit My Contacts"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuVBLINE 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeleteAllContacts 
         Caption         =   "Delete All Contacts"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuDeleteSelectedContact 
         Caption         =   "Delete Selected Contact"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuProgramProtection 
         Caption         =   "Program Protection"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rs As Recordset
Dim WS As Workspace
Dim Max As Long


Private Sub Form_Load()
Set WS = DBEngine.Workspaces(0)
    DbFile = (App.Path & "\Database\AddressBook.mdb")
    PwdString = "swordfish"
Set db = DBEngine.OpenDatabase(DbFile, False, False, ";PWD=" & PwdString)
Set rs = db.OpenRecordset("Addresses", dbOpenTable)
Max = rs.RecordCount
lblTotalContacts.Caption = Max
If rs.RecordCount = 0 Then
Exit Sub
Else
rs.MoveFirst

List1.Clear

For i = 1 To Max
    List1.AddItem rs!FullName
    rs.MoveNext
Next i
List1.ListIndex = 0
End If
ButtonsEnabled
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ButtonsEnabled
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub List1_Click()
Set WS = DBEngine.Workspaces(0)
    DbFile = (App.Path & "\Database\AddressBook.mdb")
    PwdString = "swordfish"
Set db = DBEngine.OpenDatabase(DbFile, False, False, ";PWD=" & PwdString)
Set rs = db.OpenRecordset("Select * from Addresses where FullName = '" & Trim(List1.List(List1.ListIndex)) & "'")
On Error GoTo ErrorHandler
lblName.Caption = rs("FullName")
lblNationality.Caption = rs("Nationality")
lblCountry.Caption = rs("Country")
lblBirthday.Caption = rs("Birthday")
lblCountry.Caption = rs("Country")
lblCity.Caption = rs("City")
lblAddress.Caption = rs("Address")
lblPhoneNumber.Caption = rs("PhoneNumber")
lblMobile.Caption = rs("Mobile")
lblFax.Caption = rs("Fax")
lblEMAIL.Caption = rs("E-Mail")
lblWebSite.Caption = rs("WebSite")
lblCompany.Caption = rs("Company")
Text1.Text = rs("LittleComment")
ErrorHandler:
Resume Next
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
List1_Click
PopupMenu mnuContacts
End If
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ButtonsEnabled
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show
End Sub

Private Sub mnuDeleteAllContacts_Click()
QUESTION = MsgBox("Are you sure you want to delete all your contacts ?", vbYesNo, "Warning")
If QUESTION = vbYes Then
Set WS = DBEngine.Workspaces(0)
    DbFile = (App.Path & "\Database\AddressBook.mdb")
    PwdString = "swordfish"
Set db = DBEngine.OpenDatabase(DbFile, False, False, ";PWD=" & PwdString)
Set rs = db.OpenRecordset("Addresses", dbOpenTable)
Do Until rs.RecordCount = 0
rs.MoveFirst
rs.Delete
rs.MoveNext
Loop
lblName.Caption = vbNullString
lblNationality.Caption = vbNullString
lblCountry.Caption = vbNullString
lblBirthday.Caption = vbNullString
lblCountry.Caption = vbNullString
lblCity.Caption = vbNullString
lblAddress.Caption = vbNullString
lblPhoneNumber.Caption = vbNullString
lblMobile.Caption = vbNullString
lblFax.Caption = vbNullString
lblEMAIL.Caption = vbNullString
lblWebSite.Caption = vbNullString
lblCompany.Caption = vbNullString
Text1.Text = vbNullString
List1.Clear
lblTotalContacts.Caption = 0
Else
Exit Sub
End If
End Sub

Private Sub mnuDeleteSelectedContact_Click()
QUESTION = MsgBox("Are you sure you want to delete this site link ?", vbYesNo, "Warning")
If QUESTION = vbYes Then
rs.Delete
Set rs = db.OpenRecordset("Addresses", dbOpenTable)
PutInTheListTheRecords
lblName.Caption = vbNullString
lblNationality.Caption = vbNullString
lblCountry.Caption = vbNullString
lblBirthday.Caption = vbNullString
lblCountry.Caption = vbNullString
lblCity.Caption = vbNullString
lblAddress.Caption = vbNullString
lblPhoneNumber.Caption = vbNullString
lblMobile.Caption = vbNullString
lblFax.Caption = vbNullString
lblEMAIL.Caption = vbNullString
lblWebSite.Caption = vbNullString
lblCompany.Caption = vbNullString
Text1.Text = vbNullString
Else
Exit Sub
End If
End Sub
Sub PutInTheListTheRecords()
Set WS = DBEngine.Workspaces(0)
    DbFile = (App.Path & "\Database\AddressBook.mdb")
    PwdString = "swordfish"
Set db = DBEngine.OpenDatabase(DbFile, False, False, ";PWD=" & PwdString)
Set rs = db.OpenRecordset("Addresses", dbOpenTable)
If rs.RecordCount = 0 Then
    ErrorMessageBox = MsgBox("You have 0 contacts !", vbInformation, "Information")
    List1.Clear
    Exit Sub
End If
rs.MoveLast
rs.MoveFirst
Max = rs.RecordCount
lblTotalContacts.Caption = Max
rs.MoveFirst
List1.Clear
For i = 1 To Max
    List1.AddItem rs("FullName")
    rs.MoveNext
Next i
End Sub

Private Sub mnuEditSelectedContact_Click()
If lblTotalContacts.Caption = 0 Then
Exit Sub
Else
frmEditMyContacts.Show
End If
End Sub

Sub ButtonsEnabled()
If List1 = Empty Then
mnuDeleteSelectedContact.Enabled = False
Else
mnuDeleteSelectedContact.Enabled = True
End If
End Sub

Private Sub mnuNewContact_Click()
frmNewContact.Show
End Sub

Private Sub mnuProgramProtection_Click()
frmOptions.Show
End Sub
