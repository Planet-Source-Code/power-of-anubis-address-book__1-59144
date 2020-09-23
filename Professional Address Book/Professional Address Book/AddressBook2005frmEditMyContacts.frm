VERSION 5.00
Begin VB.Form frmEditMyContacts 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit My Contacts"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6375
   Icon            =   "AddressBook2005frmEditMyContacts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   6375
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2280
      TabIndex        =   4
      Text            =   "<<Choose a letter for a fast search>>"
      Top             =   960
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
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
      Left            =   240
      MouseIcon       =   "AddressBook2005frmEditMyContacts.frx":08CA
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   7200
      Width           =   1695
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
      Height          =   6300
      Left            =   2280
      TabIndex        =   0
      Top             =   1440
      Width           =   3375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fast Search :"
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
      TabIndex        =   3
      Top             =   960
      Width           =   1860
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select the contact that you wish to edit :"
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
      TabIndex        =   1
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "frmEditMyContacts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Combo1_Click()
Dim db As Database
Dim rs As Recordset
Dim WS As Workspace
Dim Max As Long

If Combo1 = "Show All My Contacts" Then
Set WS = DBEngine.Workspaces(0)
    DbFile = (App.Path & "\Database\AddressBook.mdb")
    PwdString = "swordfish"
Set db = DBEngine.OpenDatabase(DbFile, False, False, ";PWD=" & PwdString)
Set rs = db.OpenRecordset("Addresses", dbOpenTable)
Max = rs.RecordCount
rs.MoveFirst

List1.Clear

For i = 1 To Max
    List1.AddItem rs!FullName
    rs.MoveNext
Next i
Else

Set WS = DBEngine.Workspaces(0)
    DbFile = (App.Path & "\Database\AddressBook.mdb")
    PwdString = "swordfish"
Set db = DBEngine.OpenDatabase(DbFile, False, False, ";PWD=" & PwdString)

    Set rs = db.OpenRecordset("SELECT * FROM Addresses WHERE FullName LIKE '" & Combo1.Text & "'" & "& '*'")
PutInTheListTheRecords
End If
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Combo1.AddItem "Show All My Contacts"
Combo1.AddItem "A"
Combo1.AddItem "B"
Combo1.AddItem "C"
Combo1.AddItem "D"
Combo1.AddItem "E"
Combo1.AddItem "F"
Combo1.AddItem "G"
Combo1.AddItem "H"
Combo1.AddItem "I"
Combo1.AddItem "J"
Combo1.AddItem "K"
Combo1.AddItem "L"
Combo1.AddItem "M"
Combo1.AddItem "N"
Combo1.AddItem "O"
Combo1.AddItem "P"
Combo1.AddItem "Q"
Combo1.AddItem "R"
Combo1.AddItem "S"
Combo1.AddItem "T"
Combo1.AddItem "U"
Combo1.AddItem "V"
Combo1.AddItem "W"
Combo1.AddItem "X"
Combo1.AddItem "Y"
Combo1.AddItem "Z"

Dim db As Database
Dim rs As Recordset
Dim WS As Workspace
Dim Max As Long

Set WS = DBEngine.Workspaces(0)
    DbFile = (App.Path & "\Database\AddressBook.mdb")
    PwdString = "swordfish"
Set db = DBEngine.OpenDatabase(DbFile, False, False, ";PWD=" & PwdString)
Set rs = db.OpenRecordset("Addresses", dbOpenTable)
Max = rs.RecordCount
rs.MoveFirst

List1.Clear

For i = 1 To Max
    List1.AddItem rs!FullName
    rs.MoveNext
Next i
End Sub

Sub PutInTheListTheRecords()
Dim db As Database
Dim rs As Recordset
Dim WS As Workspace
Dim Max As Long

Set WS = DBEngine.Workspaces(0)
DbFile = (App.Path & "\Database\AddressBook.mdb")
PwdString = "swordfish"
Set db = DBEngine.OpenDatabase(DbFile, False, False, ";PWD=" & PwdString)
Set rs = db.OpenRecordset("SELECT * FROM Addresses WHERE FullName LIKE '" & Combo1.Text & "'" & "& '*'")

On Error Resume Next
rs.MoveLast
rs.MoveFirst
Max = rs.RecordCount
rs.MoveFirst
List1.Clear
For i = 1 To Max
    List1.AddItem rs("FullName")
    rs.MoveNext
Next i
End Sub

Private Sub List1_Click()
Dim db As Database
Dim rs As Recordset
Dim WS As Workspace

Set WS = DBEngine.Workspaces(0)
    DbFile = (App.Path & "\Database\AddressBook.mdb")
    PwdString = "swordfish"
Set db = DBEngine.OpenDatabase(DbFile, False, False, ";PWD=" & PwdString)
Set rs = db.OpenRecordset("Select * from Addresses where FullName = '" & Trim(List1.List(List1.ListIndex)) & "'")

frmEditThisContact.txtname.Text = rs("FullName")
frmEditThisContact.txtnationality.Text = rs("Nationality")
frmEditThisContact.txtcountry.Text = rs("Country")
frmEditThisContact.txtbirthday.Text = rs("Birthday")
frmEditThisContact.txtcountry.Text = rs("Country")
frmEditThisContact.txtcity.Text = rs("City")
frmEditThisContact.txtaddress.Text = rs("Address")
frmEditThisContact.txtphonenumber.Text = rs("PhoneNumber")
frmEditThisContact.txtmobile.Text = rs("Mobile")
frmEditThisContact.txtfax.Text = rs("Fax")
frmEditThisContact.txtemail.Text = rs("E-Mail")
frmEditThisContact.txtwebsite.Text = rs("WebSite")
frmEditThisContact.txtcompany.Text = rs("Company")
frmEditThisContact.txtcomment.Text = rs("LittleComment")
frmEditThisContact.Show
Unload Me
End Sub
