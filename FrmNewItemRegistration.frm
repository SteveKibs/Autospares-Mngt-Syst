VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Frmnewitem 
   Caption         =   "New Item Form"
   ClientHeight    =   9705
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13230
   LinkTopic       =   "Form1"
   ScaleHeight     =   9705
   ScaleWidth      =   13230
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      BackColor       =   &H0080FF80&
      Height          =   975
      Left            =   1800
      TabIndex        =   27
      Top             =   6960
      Width           =   7575
      Begin VB.CommandButton Command40 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Search Item"
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox Text9 
         Height          =   495
         Left            =   720
         TabIndex        =   28
         Top             =   240
         Width           =   2895
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   360
      Top             =   2880
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Mwitidb.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Mwitidb.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "newitem"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      Height          =   1695
      Left            =   1800
      TabIndex        =   19
      Top             =   8040
      Width           =   7575
      Begin VB.CommandButton Command80 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Exit Form"
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "New Item"
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Save Item"
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Clear All"
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Delete Item"
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton cmdupdate 
         BackColor       =   &H00C0C0C0&
         Caption         =   "<< Previous Item"
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Next >>"
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1080
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2295
      Left            =   1800
      TabIndex        =   12
      Top             =   4560
      Width           =   7575
      Begin VB.TextBox Text7 
         DataField       =   "Tel No"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   3360
         TabIndex        =   15
         Top             =   1680
         Width           =   2415
      End
      Begin VB.TextBox Text6 
         DataField       =   "Address"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   3360
         TabIndex        =   14
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox Text5 
         DataField       =   "Supplier Name"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   3360
         TabIndex        =   13
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Tel No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   18
         Top             =   1680
         Width           =   2895
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   17
         Top             =   1080
         Width           =   2895
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Supplier Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   16
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4335
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      Begin VB.TextBox Text1 
         DataField       =   "dat"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   3360
         TabIndex        =   30
         Top             =   360
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         DataField       =   "item code"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   3360
         TabIndex        =   5
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox Text3 
         DataField       =   "Item Name"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   3360
         TabIndex        =   4
         Top             =   1680
         Width           =   2415
      End
      Begin VB.TextBox Text4 
         DataField       =   "Item Type"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   3360
         TabIndex        =   3
         Top             =   2400
         Width           =   2415
      End
      Begin VB.ComboBox Combo1 
         DataField       =   "Package Type"
         DataSource      =   "Adodc1"
         Height          =   315
         ItemData        =   "FrmNewItemRegistration.frx":0000
         Left            =   3360
         List            =   "FrmNewItemRegistration.frx":0016
         TabIndex        =   2
         Top             =   3120
         Width           =   2415
      End
      Begin VB.TextBox Text8 
         DataField       =   "Item Class"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   3360
         TabIndex        =   1
         Top             =   3720
         Width           =   2415
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   11
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Item Code"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   10
         Top             =   1080
         Width           =   2895
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Item Name"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   1800
         Width           =   2895
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Item Type"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   2520
         Width           =   2895
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Package Type"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   3240
         Width           =   2895
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Item Class"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   3840
         Width           =   2895
      End
   End
End
Attribute VB_Name = "Frmnewitem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub Clears()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Combo1.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text5.Text = ""

End Sub

Private Sub cmdupdate_Click()
If Adodc1.Recordset.BOF Then
Exit Sub
Else
Adodc1.Recordset.MovePrevious
End If
End Sub

Private Sub Command1_Click()
Command1.Enabled = False
Command4.Enabled = True
Command10.Enabled = True
Command5.Enabled = True
Me.Clears
End Sub

Private Sub Command10_Click()
Me.Clears
Command10.Enabled = False
Command5.Enabled = False
Command4.Enabled = True
Command1.Enabled = True
End Sub

Private Sub Command2_Click()
Data1.Recordset.Update
End Sub

Private Sub Command3_Click()
If Adodc1.Recordset.EOF Then
Exit Sub
Exit Sub
Else
Adodc1.Recordset.MoveNext
End If
End Sub

Private Sub Command4_Click()
Command4.Enabled = False

Command10.Enabled = True
Command1.Enabled = True
Command5.Enabled = True
If Text1.Text = "" Then
Exit Sub
Text1.SetFocus
Exit Sub
Else
Adodc1.Recordset.AddNew
'Adodc1.Recordset.Update
End If
End Sub

Private Sub Command5_Click()
Command1.Enabled = True
If Text1.Text = "" Then
Exit Sub
Else
Data1.Recordset.Delete
Data1.Recordset.MoveNext
If Data1.Recordset.EOF Then
If Data1.Recordset.BOF Then
Exit Sub
Else
Data1.Recordset.MoveLast
End If
End If
End If
Command5.Enabled = False

End Sub

Private Sub Command6_Click()
Data1.Recordset.MoveLast
If Data1.Recordset.EOF = True Then
MsgBox "This is the Last Record", vbInformation
End If
End Sub

Private Sub Command7_Click()

End Sub

Private Sub Command8_Click()
Data1.Recordset.MoveFirst
If Data1.Recordset.BOF Then
Exit Sub
End If
End Sub

Private Sub Command9_Click()
Command10.Enabled = True
Command5.Enabled = True

End Sub

Private Sub Command80_Click()
Unload Me
MDIForm1.Show
End Sub

Private Sub Form_Load()
Command4.Enabled = False
Command10.Enabled = False
Command5.Enabled = False
End Sub

