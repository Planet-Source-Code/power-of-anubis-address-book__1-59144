VERSION 5.00
Begin VB.Form frmNewContact 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Profile"
   ClientHeight    =   9315
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7035
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9315
   ScaleWidth      =   7035
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
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
      Left            =   5160
      MouseIcon       =   "AddressBook2005frmNewContact.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   28
      Top             =   8640
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ok"
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
      Left            =   3480
      MouseIcon       =   "AddressBook2005frmNewContact.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   27
      Top             =   8640
      Width           =   1695
   End
   Begin VB.TextBox txtcomment 
      Height          =   1455
      Left            =   1560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   26
      Top             =   7080
      Width           =   5295
   End
   Begin VB.TextBox txtcompany 
      Height          =   375
      Left            =   1560
      TabIndex        =   25
      Top             =   6360
      Width           =   5295
   End
   Begin VB.TextBox txtwebsite 
      Height          =   375
      Left            =   1560
      TabIndex        =   24
      Top             =   5880
      Width           =   5295
   End
   Begin VB.TextBox txtemail 
      Height          =   375
      Left            =   1200
      TabIndex        =   23
      Top             =   5400
      Width           =   5655
   End
   Begin VB.TextBox txtfax 
      Height          =   375
      Left            =   960
      TabIndex        =   22
      Top             =   4920
      Width           =   5895
   End
   Begin VB.TextBox txtmobile 
      Height          =   375
      Left            =   1320
      TabIndex        =   21
      Top             =   4440
      Width           =   5535
   End
   Begin VB.TextBox txtphonenumber 
      Height          =   375
      Left            =   2280
      TabIndex        =   20
      Top             =   3960
      Width           =   4575
   End
   Begin VB.TextBox txtaddress 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1440
      TabIndex        =   19
      Top             =   3240
      Width           =   5415
   End
   Begin VB.TextBox txtcity 
      Height          =   375
      Left            =   960
      TabIndex        =   18
      Top             =   2760
      Width           =   5895
   End
   Begin VB.TextBox txtcountry 
      Height          =   375
      Left            =   1440
      TabIndex        =   17
      Top             =   2280
      Width           =   5415
   End
   Begin VB.TextBox txtbirthday 
      Height          =   375
      Left            =   1440
      TabIndex        =   16
      Top             =   1800
      Width           =   5415
   End
   Begin VB.TextBox txtnationality 
      Height          =   375
      Left            =   1800
      TabIndex        =   15
      Top             =   1320
      Width           =   5055
   End
   Begin VB.TextBox txtname 
      BackColor       =   &H0080FFFF&
      Height          =   375
      Left            =   1200
      TabIndex        =   14
      Top             =   840
      Width           =   5655
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New Contact"
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
      Left            =   2400
      TabIndex        =   13
      Top             =   240
      Width           =   2235
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
      Left            =   240
      TabIndex        =   12
      Top             =   840
      Width           =   855
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
      Left            =   240
      TabIndex        =   11
      Top             =   3240
      Width           =   1155
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
      Left            =   240
      TabIndex        =   10
      Top             =   2280
      Width           =   1095
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
      Left            =   240
      TabIndex        =   9
      Top             =   3960
      Width           =   1935
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
      Left            =   240
      TabIndex        =   8
      Top             =   4440
      Width           =   945
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
      Left            =   240
      TabIndex        =   7
      Top             =   6360
      Width           =   1275
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
      Left            =   240
      TabIndex        =   6
      Top             =   5400
      Width           =   915
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
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   1425
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
      Left            =   240
      TabIndex        =   4
      Top             =   5880
      Width           =   1260
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
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Width           =   1155
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
      Left            =   240
      TabIndex        =   2
      Top             =   2760
      Width           =   615
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
      Left            =   240
      TabIndex        =   1
      Top             =   4920
      Width           =   600
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
      Left            =   240
      TabIndex        =   0
      Top             =   7080
      Width           =   1305
   End
End
Attribute VB_Name = "frmNewContact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
If txtname.Text = vbNullString Then
MsgBox "You must fill the name box!", vbExclamation, "Warning"
Exit Sub
End If
If txtnationality.Text = vbNullString Then
txtnationality.Text = "None"
End If
If txtcountry.Text = vbNullString Then
txtcountry.Text = "None"
End If
If txtcity.Text = vbNullString Then
txtcity.Text = "None"
End If
If txtbirthday.Text = vbNullString Then
txtbirthday.Text = "None"
End If
If txtaddress.Text = vbNullString Then
txtaddress.Text = "None"
End If
If txtphonenumber.Text = vbNullString Then
txtphonenumber.Text = "None"
End If
If txtmobile.Text = vbNullString Then
txtmobile.Text = "None"
End If
If txtfax.Text = vbNullString Then
txtfax.Text = "None"
End If
If txtemail.Text = vbNullString Then
txtemail.Text = "None"
End If
If txtwebsite.Text = vbNullString Then
txtwebsite.Text = "None"
End If
If txtcompany.Text = vbNullString Then
txtcompany.Text = "None"
End If
If txtcomment.Text = vbNullString Then
txtcomment.Text = "None"
End If
If txtname.Text = vbNullString Then
MsgBox "You must fill the Name text box!", vbExclamation, "Warning"
Exit Sub
Else
Dim db As Database
Dim rs As Recordset
Dim WS As Workspace
Dim Max As Long

Set WS = DBEngine.Workspaces(0)
    DbFile = (App.Path & "\Database\AddressBook.mdb")
    PwdString = "swordfish"
Set db = DBEngine.OpenDatabase(DbFile, False, False, ";PWD=" & PwdString)
Set rs = db.OpenRecordset("Addresses", dbOpenTable)
If txtname.Text = rs.Index Then
MsgBox "This contact already exists !", vbInformation, "Warning"
Else
rs.AddNew
rs("FullName") = txtname.Text
rs("Nationality") = txtnationality.Text
rs("Birthday") = txtbirthday.Text
rs("Country") = txtcountry.Text
rs("City") = txtcity.Text
rs("Address") = txtaddress.Text
rs("PhoneNumber") = txtphonenumber.Text
rs("Mobile") = txtmobile.Text
rs("Fax") = txtfax.Text
rs("E-Mail") = txtemail.Text
rs("WebSite") = txtwebsite.Text
rs("Company") = txtcompany.Text
rs("LittleComment") = txtcomment.Text
rs.Update
Max = rs.RecordCount
frmMain.lblTotalContacts.Caption = Max
rs.MoveFirst

frmMain.List1.Clear

For i = 1 To Max
    frmMain.List1.AddItem rs!FullName
    rs.MoveNext
Next i
Unload Me
End If
End If
End Sub

