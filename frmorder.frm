VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Frmorder 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Order Form"
   ClientHeight    =   9690
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14070
   LinkTopic       =   "Form3"
   ScaleHeight     =   9690
   ScaleWidth      =   14070
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   360
      Top             =   3840
      Width           =   3375
      _ExtentX        =   5953
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
      RecordSource    =   "order"
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
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FF80&
      Height          =   9255
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   11655
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFC0&
         Height          =   2535
         Left            =   120
         TabIndex        =   19
         Top             =   6600
         Width           =   11415
         Begin VB.CommandButton Command8 
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
            Left            =   6000
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   1920
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
            Left            =   3720
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   480
            Width           =   1335
         End
         Begin VB.CommandButton Command2 
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
            Left            =   5280
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   480
            Width           =   1335
         End
         Begin VB.CommandButton Command3 
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
            Left            =   6840
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   480
            Width           =   1335
         End
         Begin VB.CommandButton Command4 
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
            Left            =   8400
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   480
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
            Left            =   7440
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   1200
            Width           =   1335
         End
         Begin VB.CommandButton Command6 
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
            Left            =   4320
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   1200
            Width           =   1335
         End
         Begin VB.CommandButton Command7 
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
            Left            =   5880
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   1200
            Width           =   1335
         End
      End
      Begin VB.TextBox Text5 
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   8400
         TabIndex        =   17
         Top             =   4920
         Width           =   2415
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H0080FF80&
         Caption         =   "International"
         Height          =   255
         Left            =   5520
         TabIndex        =   16
         Top             =   5400
         Width           =   1695
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H0080FF80&
         Caption         =   "Regional"
         Height          =   255
         Left            =   5520
         TabIndex        =   15
         Top             =   4920
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H0080FF80&
         Caption         =   "Local"
         Height          =   255
         Left            =   5520
         TabIndex        =   14
         Top             =   4440
         Width           =   1695
      End
      Begin VB.TextBox Text4 
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   5400
         TabIndex        =   6
         Top             =   3600
         Width           =   2415
      End
      Begin VB.TextBox Text3 
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   5400
         TabIndex        =   5
         Top             =   2880
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   5400
         TabIndex        =   4
         Top             =   1440
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   5400
         TabIndex        =   3
         Top             =   720
         Width           =   2415
      End
      Begin VB.TextBox Text8 
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   5400
         TabIndex        =   2
         Top             =   6000
         Width           =   2415
      End
      Begin VB.TextBox Text6 
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   5400
         TabIndex        =   1
         Top             =   2160
         Width           =   2415
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ordering officer"
         BeginProperty Font 
            Name            =   "Modern No. 20"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7440
         TabIndex        =   18
         Top             =   4440
         Width           =   2895
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Order Origin"
         BeginProperty Font 
            Name            =   "Modern No. 20"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   13
         Top             =   4560
         Width           =   1695
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier"
         BeginProperty Font 
            Name            =   "Modern No. 20"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   12
         Top             =   6120
         Width           =   1695
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Delivery Date"
         BeginProperty Font 
            Name            =   "Modern No. 20"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   11
         Top             =   3720
         Width           =   2895
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Order Quantity"
         BeginProperty Font 
            Name            =   "Modern No. 20"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   10
         Top             =   3000
         Width           =   2895
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Order Number"
         BeginProperty Font 
            Name            =   "Modern No. 20"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   9
         Top             =   2280
         Width           =   2895
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Modern No. 20"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   8
         Top             =   840
         Width           =   2895
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Item Code"
         BeginProperty Font 
            Name            =   "Modern No. 20"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   7
         Top             =   1560
         Width           =   2895
      End
   End
End
Attribute VB_Name = "Frmorder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command8_Click()
Unload Me
MDIForm1.Show

End Sub
