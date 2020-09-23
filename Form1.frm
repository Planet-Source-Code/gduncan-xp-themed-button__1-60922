VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "XPBUTTON test"
   ClientHeight    =   2760
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4440
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   184
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   296
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   300
      Left            =   1800
      TabIndex        =   11
      Text            =   "someone@hotmail.com"
      Top             =   1590
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      Height          =   300
      Left            =   1800
      TabIndex        =   5
      Text            =   "Mr J. Somebody"
      Top             =   690
      Width           =   2535
   End
   Begin XPBUTTON.Duncan_XPButton btnIE 
      Height          =   375
      Index           =   2
      Left            =   2400
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FocusRect       =   0   'False
      ButtonStyle     =   1
      Enabled         =   0   'False
      Picture         =   "Form1.frx":058A
      PictureAlignment=   37
   End
   Begin XPBUTTON.Duncan_XPButton btnIE 
      Height          =   375
      Index           =   1
      Left            =   1200
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FocusRect       =   0   'False
      ButtonStyle     =   1
      Caption         =   "IE Button 2"
   End
   Begin XPBUTTON.Duncan_XPButton Duncan_XPButton5 
      Height          =   375
      Left            =   90
      TabIndex        =   6
      Top             =   1080
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FocusRect       =   0   'False
      Caption         =   "Home Email"
   End
   Begin XPBUTTON.Duncan_XPButton btnIE 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FocusRect       =   0   'False
      ButtonStyle     =   1
      Caption         =   "IE Button 1"
   End
   Begin XPBUTTON.Duncan_XPButton Duncan_XPButton3 
      Height          =   375
      Left            =   1380
      TabIndex        =   7
      Top             =   1080
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Marlett"
         Size            =   8.25
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FocusRect       =   0   'False
      Caption         =   "6"
   End
   Begin XPBUTTON.Duncan_XPButton Duncan_XPButton2 
      Cancel          =   -1  'True
      Height          =   615
      Left            =   1800
      TabIndex        =   13
      Top             =   2040
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Cancel"
      Picture         =   "Form1.frx":0924
      PictureAlignment=   38
      PicturePadding  =   -2
   End
   Begin XPBUTTON.Duncan_XPButton Duncan_XPButton1 
      Default         =   -1  'True
      Height          =   615
      Left            =   3120
      TabIndex        =   12
      Top             =   2040
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Save"
      CaptionAlignment=   41
      Picture         =   "Form1.frx":0CBE
      PictureAlignment=   33
      PicturePadding  =   2
      CaptionPadding  =   2
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   1800
      TabIndex        =   8
      Text            =   "jsomebody@somedomain.com"
      Top             =   1110
      Width           =   2535
   End
   Begin XPBUTTON.Duncan_XPButton btnIE 
      Height          =   375
      Index           =   3
      Left            =   2880
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FocusRect       =   0   'False
      ButtonStyle     =   1
      Caption         =   "IE Button 4"
   End
   Begin XPBUTTON.Duncan_XPButton Duncan_XPButton4 
      Height          =   375
      Left            =   90
      TabIndex        =   9
      Top             =   1560
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FocusRect       =   0   'False
      Caption         =   "Assistants Email"
   End
   Begin XPBUTTON.Duncan_XPButton Duncan_XPButton6 
      Height          =   375
      Left            =   1380
      TabIndex        =   10
      Top             =   1560
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Marlett"
         Size            =   8.25
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FocusRect       =   0   'False
      Caption         =   "6"
   End
   Begin XPBUTTON.Duncan_XPButton btnIE 
      Height          =   375
      Index           =   4
      Left            =   3960
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FocusRect       =   0   'False
      ButtonStyle     =   1
      Picture         =   "Form1.frx":1012
      PictureAlignment=   37
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Contacts Name:"
      Height          =   255
      Left            =   480
      TabIndex        =   14
      Top             =   720
      Width           =   1215
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuItem 
         Caption         =   "Home Email"
         Index           =   0
      End
      Begin VB.Menu mnuItem 
         Caption         =   "Work Email"
         Index           =   1
      End
      Begin VB.Menu mnuItem 
         Caption         =   "Assistants Email"
         Index           =   2
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Item As String

Private Sub Duncan_XPButton1_Click()
    Debug.Print "clicked btn 1 - default"
End Sub

Private Sub Duncan_XPButton2_Click()
    Debug.Print "clicked btn 2 - cancel"
End Sub

Private Sub Duncan_XPButton3_Click()
    PopupMenu mnuPopup, , Duncan_XPButton3.left, Duncan_XPButton3.top + Duncan_XPButton3.Height
    If Len(m_Item) > 0 Then
        'they selected an item
        Duncan_XPButton5.Caption = m_Item
        m_Item = ""
    End If
End Sub

Private Sub btnIE_Click(Index As Integer)
    Debug.Print "IE button " & Index + 1
End Sub

Private Sub Duncan_XPButton4_Click()
    MsgBox "send mail to " & Text1
End Sub

Private Sub Duncan_XPButton5_Click()
    MsgBox "send mail to " & Text1
End Sub

Private Sub Duncan_XPButton6_Click()
    PopupMenu mnuPopup, , Duncan_XPButton3.left, Duncan_XPButton6.top + Duncan_XPButton6.Height
    If Len(m_Item) > 0 Then
        'they selected an item
        Duncan_XPButton4.Caption = m_Item
        m_Item = ""
    End If
End Sub

Private Sub mnuItem_Click(Index As Integer)
    m_Item = mnuItem(Index).Caption
    
End Sub
