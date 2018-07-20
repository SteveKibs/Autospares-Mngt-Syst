VERSION 5.00
Begin VB.Form Frmsales 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Sales Form"
   ClientHeight    =   9360
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11640
   LinkTopic       =   "Form2"
   ScaleHeight     =   9360
   ScaleWidth      =   11640
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FF80&
      Height          =   6135
      Left            =   1320
      TabIndex        =   9
      Top             =   600
      Width           =   10575
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   5760
         TabIndex        =   16
         Top             =   3480
         Width           =   2415
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   5760
         TabIndex        =   15
         Top             =   2760
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   5760
         TabIndex        =   14
         Top             =   1320
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   5760
         TabIndex        =   13
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Index           =   0
         Left            =   5760
         TabIndex        =   12
         Top             =   5160
         Width           =   2415
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   5760
         TabIndex        =   11
         Top             =   4320
         Width           =   2415
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   5760
         TabIndex        =   10
         Top             =   2040
         Width           =   2415
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Sale"
         BeginProperty Font 
            Name            =   "Modern No. 20"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   375
         Left            =   2760
         TabIndex        =   23
         Top             =   4440
         Width           =   2895
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Item Price Per"
         BeginProperty Font 
            Name            =   "Modern No. 20"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   375
         Left            =   2760
         TabIndex        =   22
         Top             =   3600
         Width           =   2895
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
         BeginProperty Font 
            Name            =   "Modern No. 20"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   375
         Left            =   2760
         TabIndex        =   21
         Top             =   2880
         Width           =   2895
      End
      Begin VB.Label Label2 
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
         ForeColor       =   &H00004000&
         Height          =   375
         Left            =   2760
         TabIndex        =   20
         Top             =   1440
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
         ForeColor       =   &H00004000&
         Height          =   375
         Left            =   2760
         TabIndex        =   19
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Officer in Charge"
         BeginProperty Font 
            Name            =   "Modern No. 20"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   375
         Left            =   2760
         TabIndex        =   18
         Top             =   5280
         Width           =   2895
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Item Name"
         BeginProperty Font 
            Name            =   "Modern No. 20"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   375
         Left            =   2760
         TabIndex        =   17
         Top             =   2160
         Width           =   2895
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
      Height          =   2535
      Left            =   1320
      TabIndex        =   0
      Top             =   6720
      Width           =   10575
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
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1200
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
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1200
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
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   480
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
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   480
         Width           =   1335
      End
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
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1920
         Width           =   1335
      End
   End
End
Attribute VB_Name = "Frmsales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command8_Click()
Unload Me
MDIForm1.Show

End Sub
