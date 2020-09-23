VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form FrmMain 
   BackColor       =   &H00F9E2D1&
   Caption         =   "                                                                                    PSC Zip File Reader"
   ClientHeight    =   9480
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   12750
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "FrmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9480
   ScaleWidth      =   12750
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Authors"
      Height          =   3645
      Left            =   4395
      TabIndex        =   41
      Top             =   4845
      Visible         =   0   'False
      Width           =   8100
      Begin VB.ListBox lstAuthor2 
         Height          =   2985
         Left            =   4275
         TabIndex        =   46
         Top             =   555
         Width           =   3690
      End
      Begin VB.ListBox lstAuthors 
         Height          =   2985
         Left            =   135
         TabIndex        =   43
         Top             =   555
         Width           =   4050
      End
      Begin VB.TextBox txtSearch 
         Height          =   255
         Left            =   150
         TabIndex        =   42
         Top             =   225
         Width           =   1755
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "All the projects in database           by this author"
         Height          =   375
         Left            =   5430
         TabIndex        =   47
         Top             =   150
         Width           =   2010
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Search (type in name or click name in list)"
         Height          =   240
         Left            =   2040
         TabIndex        =   44
         Top             =   240
         Width           =   3135
      End
   End
   Begin RichTextLib.RichTextBox rtb1 
      Height          =   4515
      Left            =   45
      TabIndex        =   10
      Top             =   4575
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   7964
      _Version        =   393217
      BackColor       =   14416636
      ScrollBars      =   2
      TextRTF         =   $"FrmMain.frx":08CA
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFC&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4350
      Left            =   6720
      TabIndex        =   9
      Top             =   4575
      Width           =   5955
   End
   Begin VB.TextBox txtRightStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Height          =   285
      Left            =   6735
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   9135
      Width           =   5940
   End
   Begin VB.TextBox txtLeftStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Height          =   285
      Left            =   45
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   9135
      Width           =   6645
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFC&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4440
      Left            =   8565
      TabIndex        =   5
      Top             =   30
      Width           =   4110
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBFAFC&
      Caption         =   "Viewer/ Entry"
      Height          =   4500
      Left            =   45
      TabIndex        =   1
      Top             =   30
      Width           =   6645
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   525
         Left            =   4050
         ScaleHeight     =   495
         ScaleWidth      =   2460
         TabIndex        =   39
         Top             =   2010
         Width           =   2490
         Begin VB.TextBox txtComments 
            Appearance      =   0  'Flat
            Height          =   525
            Left            =   -15
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   40
            Top             =   -15
            Width           =   2475
         End
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4050
         ScaleHeight     =   255
         ScaleWidth      =   2460
         TabIndex        =   37
         Top             =   990
         Width           =   2490
         Begin VB.TextBox txtAuthorName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00CBFAFD&
            Height          =   285
            Left            =   -15
            TabIndex        =   38
            Top             =   -15
            Width           =   2490
         End
      End
      Begin Project1.TextEffects TextEffects1 
         Height          =   480
         Left            =   1860
         TabIndex        =   31
         Top             =   3960
         Width           =   4725
         _ExtentX        =   8334
         _ExtentY        =   847
         TextStyle       =   2
         Text            =   "PSC ZIP FILE READER"
         TextColor       =   8438015
      End
      Begin Project1.ThemedButton cmdCancelEF 
         Height          =   555
         Left            =   5460
         TabIndex        =   28
         Top             =   2625
         Visible         =   0   'False
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   979
         BackColor       =   16777215
         Caption         =   "Cancel"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "FrmMain.frx":0946
         ShowFocusRect   =   0   'False
      End
      Begin Project1.ThemedButton cmdCancelNF 
         Height          =   555
         Left            =   5460
         TabIndex        =   27
         Top             =   2610
         Visible         =   0   'False
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   979
         BackColor       =   16777215
         Caption         =   "Cancel"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "FrmMain.frx":0EE0
         ShowFocusRect   =   0   'False
      End
      Begin Project1.ThemedButton cmdBrowse 
         Height          =   555
         Left            =   3990
         TabIndex        =   26
         Top             =   2625
         Visible         =   0   'False
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   979
         BackColor       =   16777215
         Caption         =   "Browse For Zip"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "FrmMain.frx":147A
         ShowFocusRect   =   0   'False
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4050
         ScaleHeight     =   255
         ScaleWidth      =   2460
         TabIndex        =   18
         Top             =   480
         Width           =   2490
         Begin VB.TextBox txtZipName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00CBFAFD&
            Height          =   285
            Left            =   -15
            TabIndex        =   19
            Top             =   -15
            Width           =   2490
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   195
         ScaleHeight     =   345
         ScaleWidth      =   1440
         TabIndex        =   16
         Top             =   4005
         Width           =   1470
         Begin Project1.ThemedButton cmdRight 
            Height          =   285
            Left            =   990
            TabIndex        =   30
            Top             =   30
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   503
            BackColor       =   0
            Caption         =   ">>"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "FrmMain.frx":1A14
            ShowFocusRect   =   0   'False
         End
         Begin Project1.ThemedButton cmdLeft 
            Height          =   285
            Left            =   30
            TabIndex        =   29
            Top             =   30
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   503
            BackColor       =   0
            Caption         =   "<<"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "FrmMain.frx":1FAE
            ShowFocusRect   =   0   'False
         End
         Begin VB.TextBox txtpost 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00DBFAFC&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   480
            TabIndex        =   17
            Text            =   "1"
            Top             =   30
            Width           =   495
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4050
         ScaleHeight     =   255
         ScaleWidth      =   2460
         TabIndex        =   14
         Top             =   1515
         Width           =   2490
         Begin VB.TextBox txtPicName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00CBFAFD&
            Height          =   285
            Left            =   -15
            TabIndex        =   15
            Top             =   -15
            Width           =   2490
         End
      End
      Begin VB.PictureBox picBuffer 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00DBFAFC&
         BorderStyle     =   0  'None
         Height          =   3705
         Left            =   105
         ScaleHeight     =   247
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   255
         TabIndex        =   6
         Top             =   210
         Width           =   3825
      End
      Begin VB.PictureBox picMain 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         Height          =   3705
         Left            =   105
         ScaleHeight     =   247
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   255
         TabIndex        =   4
         Top             =   210
         Visible         =   0   'False
         Width           =   3825
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Comments"
         ForeColor       =   &H00C000C0&
         Height          =   195
         Left            =   4080
         TabIndex        =   36
         Top             =   1815
         Width           =   825
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "4. Save"
         Height          =   180
         Left            =   4275
         TabIndex        =   35
         Top             =   3750
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "3. Add picture and comments."
         Height          =   225
         Left            =   4275
         TabIndex        =   34
         Top             =   3570
         Visible         =   0   'False
         Width           =   2190
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "2. Add authors name"
         Height          =   240
         Left            =   4275
         TabIndex        =   33
         Top             =   3390
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "1. Browse for zip"
         Height          =   270
         Left            =   4275
         TabIndex        =   32
         Top             =   3210
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.Shape Shape3 
         Height          =   2370
         Left            =   4005
         Top             =   225
         Width           =   2580
      End
      Begin VB.Shape Shape1 
         Height          =   4500
         Left            =   0
         Top             =   0
         Width           =   6645
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Picture Name in Photo Folder"
         ForeColor       =   &H000040C0&
         Height          =   225
         Left            =   4065
         TabIndex        =   13
         Top             =   1290
         Width           =   2145
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Zip File"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   4050
         TabIndex        =   3
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Author"
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   4050
         TabIndex        =   2
         Top             =   780
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00F9E2D1&
      Caption         =   "Control Panel"
      Height          =   4500
      Left            =   6720
      TabIndex        =   0
      Top             =   30
      Width           =   1815
      Begin Project1.ThemedButton cmdShowAuthors 
         Height          =   480
         Left            =   105
         TabIndex        =   45
         Top             =   1650
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   847
         Caption         =   "Authors List"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "FrmMain.frx":2548
         Picture         =   "FrmMain.frx":2AE2
         PictureAlign    =   1
         PictureSize     =   0
      End
      Begin Project1.ThemedButton cmdLaunch 
         Height          =   750
         Left            =   75
         TabIndex        =   25
         Top             =   3675
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   1323
         BackColor       =   16376529
         Caption         =   "Launch A Selected File From Below"
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483633
         MouseIcon       =   "FrmMain.frx":3734
         Picture         =   "FrmMain.frx":3CCE
         PictureAlign    =   3
         PictureSize     =   0
         ShowFocusRect   =   0   'False
      End
      Begin Project1.ThemedButton cmdQuit 
         Height          =   495
         Left            =   90
         TabIndex        =   24
         Top             =   3105
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   873
         BackColor       =   16376529
         Caption         =   "Exit"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   255
         MouseIcon       =   "FrmMain.frx":41C3
         Picture         =   "FrmMain.frx":475D
         PictureAlign    =   1
         PictureSize     =   0
         ShowFocusRect   =   0   'False
      End
      Begin Project1.ThemedButton cmdDeleteAll 
         Height          =   480
         Left            =   105
         TabIndex        =   23
         Top             =   2145
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   847
         BackColor       =   16376529
         Caption         =   "Delete All"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   255
         MouseIcon       =   "FrmMain.frx":53AF
         Picture         =   "FrmMain.frx":5949
         PictureAlign    =   1
         PictureSize     =   0
         ShowFocusRect   =   0   'False
      End
      Begin Project1.ThemedButton cmdDeleteCont 
         Height          =   480
         Left            =   105
         TabIndex        =   22
         Top             =   1155
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   847
         BackColor       =   16376529
         Caption         =   "Delete File"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16576
         MouseIcon       =   "FrmMain.frx":659B
         Picture         =   "FrmMain.frx":6B35
         PictureAlign    =   1
         PictureSize     =   0
         ShowFocusRect   =   0   'False
      End
      Begin Project1.ThemedButton cmdEditCont 
         Height          =   465
         Left            =   105
         TabIndex        =   21
         Top             =   675
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   820
         BackColor       =   16376529
         Caption         =   "Edit File"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   49152
         MouseIcon       =   "FrmMain.frx":7787
         Picture         =   "FrmMain.frx":7D21
         PictureAlign    =   1
         PictureSize     =   0
         ShowFocusRect   =   0   'False
      End
      Begin Project1.ThemedButton cmdNewCont 
         Height          =   450
         Left            =   105
         TabIndex        =   20
         Top             =   210
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   794
         BackColor       =   16376529
         Caption         =   "New File"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
         MouseIcon       =   "FrmMain.frx":8973
         Picture         =   "FrmMain.frx":8F0D
         PictureAlign    =   1
         PictureSize     =   0
         ShowFocusRect   =   0   'False
      End
      Begin VB.Shape Shape2 
         Height          =   4500
         Left            =   0
         Top             =   0
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "single click to select >"
         Height          =   210
         Left            =   165
         TabIndex        =   12
         Top             =   2655
         Width           =   1605
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "dbl click to show files >"
         Height          =   315
         Left            =   75
         TabIndex        =   11
         Top             =   2880
         Width           =   1695
      End
   End
   Begin VB.Menu mnuMnu 
      Caption         =   "Menu"
      Begin VB.Menu mnuNewFile 
         Caption         =   "New File"
      End
      Begin VB.Menu mnuEditFile 
         Caption         =   "Edit File"
      End
      Begin VB.Menu mnuDeleteFile 
         Caption         =   "Delete File"
      End
      Begin VB.Menu mnuDeleteAll 
         Caption         =   "Delete All"
      End
      Begin VB.Menu dash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAuthorsList 
         Caption         =   "Authors List"
      End
      Begin VB.Menu dash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuKeepPics 
         Caption         =   "Keep Photos in Folder"
         Checked         =   -1  'True
      End
      Begin VB.Menu dash3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'by Ken Foster Dec 2009
'MS Access was used to create database table
'The Photos folder is for convinence. If you should loose picture
'in database, you won't have to search all over to find it again.
'There is an option to save or not save the photos in the folder. See menu
'Make sure zip file is the name you want.(rename it if necessary, will ask before saving)
'As you save, the picture(if option is checked) and zip are copied into their respective folder'
'so you do not have to manually do it.
'Also the new picture name is CHANGED to match the zip files name.
'I also chose to have the zipped files in a folder and not the database,
'in case you want to have access the the zip file.
'If you are viewing the code of a zip file, it is temporaryly unzipped
'in the tempunzip folder and is cleared when you choose another or exit
'this program.

Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const LB_SETTABSTOPS = &H192
Private Const LB_ITEMFROMPOINT As Long = &H1A9

Dim editing As Boolean                               'in edit mode
Dim Seld As Boolean                                  'add picture to database selected
Dim PhotosInFolder As Integer                      'store pictures in Photo Folder
Dim KeepCurrentPic As Boolean                        'keep current picture in database
Dim sBuffer As String                                ' store last selection
Private bzip As CGUnzipFiles
Dim afilepath As String                              'path to zip file
Dim Shadow As clsShadow

Private Sub Form_Load()
    ReDim TabStop(0) As Long
    
     ' Set Shadow to form
    Set Shadow = New clsShadow
    Call Shadow.Shadow(Me)
    Shadow.Color = vbBlack
    Shadow.Depth = 10
    Shadow.Transparency = 180
    
    'set tabs for lstAuthors and lstAuthor2 listboxes
    TabStop(0) = 70
    Call SendMessage(lstAuthors.hwnd, LB_SETTABSTOPS, 0&, ByVal 0&)
    Call SendMessage(lstAuthors.hwnd, LB_SETTABSTOPS, 1, TabStop(0))
    lstAuthors.Refresh
    Call SendMessage(lstAuthor2.hwnd, LB_SETTABSTOPS, 0&, ByVal 0&)
    Call SendMessage(lstAuthor2.hwnd, LB_SETTABSTOPS, 1, TabStop(0))
    lstAuthor2.Refresh
    
    KeepCurrentPic = True
    Seld = False
    PhotosInFolder = Int(LoadText(App.Path & "\Options.txt"))       ' keep pictures in Photos folder or not
    If PhotosInFolder = 1 Then
       mnuKeepPics.Checked = True
    Else
       mnuKeepPics.Checked = False
    End If
    postnr = 1
    If Dir(App.Path & "\klf.mdb") <> "" Then GoTo continue
    CreateNewDB "klf.mdb"
continue:
    Set bzip = New CGUnzipFiles
    InitKeyWords
    Set Db = OpenDatabase(App.Path & "\klf.mdb")
    
    Set Rs = Db.OpenRecordset("data")
    
    If Rs.RecordCount > 0 Then
        
        Rs.Close: Set Rs = Nothing: Db.Close: Set Db = Nothing
        visapost (1)
        Open App.Path & "\klf.mdb" For Binary As #1
        g = LOF(1)
        Close #1
        LoadListbox
        List1.ListIndex = 0
    Else
        txtLeftStatus.Text = "There are no files in the database."
        cmdDeleteAll.Enabled = False
        txtpost = "0"
        cmdDeleteCont.Enabled = False
        cmdEditCont.Enabled = False
    End If
    Open App.Path & "\klf.mdb" For Binary As #1
    g = LOF(1)
    Close #1
    txtRightStatus.Text = "Size of the database : " & Format(g, "###,###,###,##0") & " k"
    sBuffer = App.Path
    LoadAuthors
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    KillFolder App.Path & "\tempUnzip\", True  'dump any files that are still in tempUnzip folder
    DoEvents
    Set bzip = Nothing
    Unload Me
End Sub

Private Sub Form_Resize()
    txtLeftStatus.TOP = Me.Height - 1150
    txtRightStatus.TOP = Me.Height - 1150
    rtb1.Height = txtLeftStatus.TOP - 4670
    List2.Height = rtb1.Height
    If Me.Width > 12870 Then Me.Width = 12870
    List1.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    KillFolder App.Path & "\tempUnzip\", True  'dump any files that are still in tempUnzip folder
    DoEvents
    Set bzip = Nothing
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 27 Then    'esc key
    
    If editing = True Then
        picMain.Cls
        
        Set Db = OpenDatabase(App.Path & "\klf.mdb")
        Set Rs = Db.OpenRecordset("data")
        cmdNewCont.Caption = "&New File"
        cmdEditCont.Caption = "&Edit File"
        If Rs.RecordCount > 0 Then
            cmdDeleteCont.Enabled = True
            cmdEditCont.Enabled = True
        End If
        Rs.Close: Set Rs = Nothing: Db.Close: Set Db = Nothing
        editing = False
        visapost postnr
        
    End If
End If
End Sub

Private Sub cmdBrowse_Click()
    Dim resp As String
    
    afilepath = GetFolder(Me.hwnd, sBuffer, "Please Select a Zip File", True, False)
    If afilepath = "" Then Exit Sub   'cancel pressed
    sBuffer = afilepath
    
    txtZipName.Text = Mid$(afilepath, InStrRev(afilepath, "\") + 1, Len(afilepath))
    
    resp = MsgBox("Is zip name correct? If not, correct and press Enter bar.", vbYesNo, "Correct Zip Name")
    If resp = 7 Then Exit Sub
    
    If Dir(App.Path & "\ZippedFiles\" & txtZipName.Text) <> "" Then ' it exists already
    resp = MsgBox("Zip filename already exists. Do you want to replace it?", vbYesNo, "Zip FileName already exists")
    If resp = 7 Then      'no
    txtZipName.Text = ""
    txtAuthorName.Text = ""
    cmdNewCont_Click
    Exit Sub
End If
End If
DeleteFile App.Path & "\ZippedFiles\" & txtZipName.Text            'delete old zip in folder
FileCopy afilepath, App.Path & "\ZippedFiles\" & txtZipName.Text   'copy new zip file into Folder
End Sub

Private Sub cmdCancelEF_Click()
    KeepCurrentPic = True
    Seld = False
    cmdEditCont_Click
    cmdCancelEF.Visible = False
    'disable buttons
    Picture3.Enabled = False
    Picture4.Enabled = False
    Picture5.Enabled = False
End Sub

Private Sub cmdCancelNF_Click()
    txtZipName.Text = ""
    cmdNewCont_Click
    cmdCancelNF.Visible = False
    'instruction labels
    Label6.Visible = False
    Label7.Visible = False
    Label8.Visible = False
    Label9.Visible = False
    'disable buttons
    Picture3.Enabled = False
    Picture4.Enabled = False
    Picture5.Enabled = False
End Sub

Private Sub cmdNewCont_Click()
    
    If editing = False Then
        cmdBrowse.Visible = True
        cmdLeft.Enabled = False
        cmdRight.Enabled = False
        txtpost.Locked = True
        picBuffer.Visible = False
        picMain.Visible = True
        picBuffer.Picture = LoadPicture()
        picMain.Picture = LoadPicture()
        picMain.Cls
        picMain.Print
        picMain.Print "      Double click here to add a picture"
        editing = True
        KeepCurrentPic = False
        Seld = True
        cmdDeleteCont.Enabled = False
        cmdEditCont.Enabled = False
        cmdDeleteAll.Enabled = False
        'disables textboxes but they do not show gray
        Picture3.Enabled = True
        Picture4.Enabled = True
        Picture5.Enabled = True
        
        txtZipName.Text = ""
        txtAuthorName.Text = ""
        txtPicName.Text = ""
        txtComments.Text = ""
        'instruction labels
        Label6.Visible = True
        Label7.Visible = True
        Label8.Visible = True
        Label9.Visible = True
        
        cmdCancelNF.Visible = True
        cmdNewCont.Caption = "&Save File"
        picMain.Visible = True
    Else
        picMain.Cls
        
        If txtAuthorName.Text = "" Then txtAuthorName.Text = "No Author name available."
        If txtComments.Text = "" Then txtComments.Text = "none"
        If txtZipName.Text <> "" Then savetodb
        editing = False
        Seld = False
        KeepCurrentPic = True
        cmdNewCont.Caption = "&New File"
        picMain.Visible = False
        cmdBrowse.Visible = False
        cmdLeft.Enabled = True
        cmdRight.Enabled = True
        txtpost.Locked = False
        cmdDeleteCont.Enabled = True
        cmdEditCont.Enabled = True
        cmdDeleteAll.Enabled = True
        cmdCancelNF.Visible = False
        'instruction labels
        Label6.Visible = False
        Label7.Visible = False
        Label8.Visible = False
        Label9.Visible = False
        
        LoadListbox
        If List1.ListCount = 0 Then Exit Sub
        List1.ListIndex = Val(txtpost.Text) - 1
        List1_Click
    End If
    'if any changes were made then reload list
    LoadAuthors
    lstAuthor2.Clear
End Sub

Private Sub cmdDeleteCont_Click()
    Dim resp As String
    Dim tmp As Integer
    
    If List1.ListIndex = -1 Then Exit Sub
    resp = MsgBox("Are you sure you want to delete the selected file?", vbYesNo, "PSC Zip Viewer")
    If resp = 7 Then Exit Sub          'no

    DeleteFile App.Path & "\ZippedFiles\" & txtZipName.Text   'delete zip from ZippedFiles folder
    DeleteFile App.Path & "\Photos\" & txtPicName.Text        'delete picture from Photo folder
    
    Set Db = OpenDatabase(App.Path & "\klf.mdb")
    Set Rs = Db.OpenRecordset("data")
    Rs.Move postnr - 1
    Rs.Delete
    postnr = Rs.RecordCount
    txtpost = postnr
    
    If Rs.RecordCount > 0 Then
        tmp = Rs.RecordCount
        Rs.Close: Set Rs = Nothing: Db.Close: Set Db = Nothing
        visapost tmp
        Open App.Path & "\klf.mdb" For Binary As #1
        g = LOF(1)
        Close #1
        
        txtRightStatus.Text = "Size of the database : " & Format(g, "###,###,###,##0") & " k"
    Else
        txtZipName.Text = ""
        txtAuthorName.Text = ""
        txtPicName.Text = ""
        picBuffer.Picture = LoadPicture()
        txtComments.Text = ""
        'disables textboxes but they do not show gray
        Picture3.Enabled = False
        Picture4.Enabled = False
        Picture5.Enabled = False
        
        txtLeftStatus.Text = "There are no files in the database."
        cmdDeleteAll.Enabled = False
        txtRightStatus.Text = "Size of the database : " & Format(g, "###,###,###,##0") & " k"
        cmdDeleteCont.Enabled = False
        cmdEditCont.Enabled = False
    End If
    If List1.ListCount <> 0 Then LoadListbox
    'if any changes were made then reload list
    LoadAuthors
    lstAuthor2.Clear
End Sub

Private Sub cmdLaunch_Click()
    Shell (List2.Text)
End Sub

Private Sub cmdLeft_Click()
    Set Db = OpenDatabase(App.Path & "\klf.mdb")
    Set Rs = Db.OpenRecordset("data")
    If Rs.RecordCount > 0 Then
        postnr = postnr - 1
        If postnr < 1 Then postnr = 1
        txtpost = Str(postnr)
        visapost postnr
    End If
    List1.ListIndex = postnr - 1
End Sub

Private Sub cmdRight_Click()
    Set Db = OpenDatabase(App.Path & "\klf.mdb")
    Set Rs = Db.OpenRecordset("data")
    If Rs.RecordCount > 0 Then
        postnr = postnr + 1
        If postnr > Rs.RecordCount Then postnr = Rs.RecordCount
        Rs.Close: Set Rs = Nothing: Db.Close: Set Db = Nothing
        visapost postnr
        List1.ListIndex = postnr - 1
    End If
End Sub

Private Sub cmdEditCont_Click()
    If txtZipName.Text = "" Then GoTo here1
    If editing = False Then
        cmdLeft.Enabled = False
        cmdRight.Enabled = False
        txtpost.Locked = True
        If picBuffer.Picture = 0 Then
            picBuffer.Visible = False
            picMain.Visible = True
            picMain.Picture = LoadPicture()
            picMain.Cls
            picMain.Print
            picMain.Print "      Double click here to add a picture"
            picMain.Print
            picMain.Print
            picMain.Print "      IGNORE ABOVE IF IN EDIT MODE."
            picMain.Print "   Unless you really want a different picture."
        End If
        editing = True
        cmdNewCont.Enabled = False
        cmdDeleteCont.Enabled = False
        cmdDeleteAll.Enabled = False
        cmdEditCont.Caption = "&Save File"
        'disables textboxes but they do not show gray
        Picture3.Enabled = True
        Picture4.Enabled = True
        Picture5.Enabled = True
        
        cmdCancelEF.Visible = True
        LoadListbox
    Else
        editing = False
        cmdLeft.Enabled = True
        cmdRight.Enabled = True
        txtpost.Locked = False
        picMain.Cls
        If txtComments.Text = "" Then txtComments.Text = "none"
        updatedb (postnr)
here1:
        cmdNewCont.Enabled = True
        cmdDeleteCont.Enabled = True
        cmdDeleteAll.Enabled = True
        cmdEditCont.Caption = "&Edit File"
        picMain.Visible = False
        List1.ListIndex = Val(txtpost.Text) - 1
        List1_Click
        KeepCurrentPic = True
        Seld = False
        cmdCancelEF.Visible = False
    End If
    'if any changes were made then reload list
    LoadAuthors
    lstAuthor2.Clear
End Sub

Private Sub cmdQuit_Click()
    Dim resp As String
    If editing = True Then
        resp = MsgBox("Are you sure you want to quit?" & vbCrLf & "The changes you have made in the current contact will not be saved.", vbYesNo, "Contacts 1.0")
        If resp = 7 Then Exit Sub
    End If
    Set Db = OpenDatabase(App.Path & "\klf.mdb")
    Set Rs = Db.OpenRecordset("data")
    If Rs.RecordCount > 0 Then
        Rs.Close: Set Rs = Nothing: Db.Close: Set Db = Nothing
        
        Call CompactDatabase(App.Path & "\klf.mdb", App.Path & "\compact.mdb")
        Kill App.Path & "\klf.mdb"
        Call FileCopy(App.Path & "\compact.mdb", App.Path & "\klf.mdb")
        Kill App.Path & "\compact.mdb"
    End If
    Unload Me
    End
End Sub

Private Sub cmdDeleteAll_Click()
    Dim resp As String
    resp = MsgBox("Are you sure you want to delete all the contacts in the database?", vbYesNo, "PSC Zip Viewer")
    If resp = 7 Then Exit Sub
    Call CreateNewDB("klf.mdb")
    Form_Load
    
    txtZipName.Text = ""
    txtAuthorName.Text = ""
    txtPicName.Text = ""
    picBuffer.Picture = LoadPicture()
    txtComments.Text = ""
    List1.Clear
End Sub

Private Sub cmdShowAuthors_Click()
    Frame3.Visible = Not Frame3.Visible
    txtSearch.Text = ""
    LoadAuthorList
    If Frame3.Visible = True Then txtSearch.SetFocus
End Sub
    
Public Sub savetodb()
    Dim filnam As String
    Dim ext As String
    Dim strFromFile As String
    Dim lngFileSize As Long
    Dim filenum As Integer
    
    Set Db = OpenDatabase(App.Path & "\klf.mdb")
    'Open the RecordSet to get the Categories List
    Set Rs = Db.OpenRecordset("data")
    On Error Resume Next
    Rs.AddNew
    
    Rs.Fields(0) = Trim(txtZipName.Text)
    Rs.Fields(1) = Trim(txtAuthorName.Text)
    Rs.Fields(4) = Trim(txtComments.Text)
    If picBuffer.Picture > "" Then
        If Seld = True Then
            'rename picture with zip name so they are the same
            ext = Mid$(cmndlg.FileName, InStrRev(cmndlg.FileName, ".", , vbTextCompare) + 1)
            If Len(ext) = 3 Then filnam = LEFT$(txtZipName.Text, Len(txtZipName.Text) - 3)
            If Len(ext) = 4 Then filnam = LEFT$(txtZipName.Text, Len(txtZipName.Text) - 4)
            
            Call SavePicture(picBuffer.Picture, App.Path & "\tmpfile")
            If PhotosInFolder = 1 Then FileCopy cmndlg.FileName, App.Path & "\Photos\" & filnam & ext    'copy picture into Photo folder
            Rs.Fields(2) = filnam & ext
            Seld = False
            KeepCurrentPic = True
        End If
      
        filenum = FreeFile
        lngFileSize = FileLen(App.Path & "\tmpfile")
        strFromFile = String(lngFileSize, " ")
        Open App.Path & "\tmpfile" For Binary As filenum
        Get filenum, , strFromFile
        Close filenum
        
        Rs.Fields(3) = strFromFile
        Kill App.Path & "\tmpfile"
    End If
    Rs.Update
    Rs.MoveLast
    
    txtLeftStatus.Text = "There are" & Str(Rs.RecordCount) & " files in the database. Showing file " & Str(Rs.RecordCount) & "."
    txtpost.Text = Str(Rs.RecordCount)
    postnr = Rs.RecordCount
    Rs.Close: Set Rs = Nothing: Db.Close: Set Db = Nothing
    
    Open App.Path & "\klf.mdb" For Binary As #1
    g = LOF(1)
    Close #1
    txtRightStatus.Text = "Size of the database : " & Format(g, "###,###,###,##0") & " k"
    visapost (postnr)
    'disables textboxes but they do not show gray
    Picture3.Enabled = False
    Picture4.Enabled = False
    Picture5.Enabled = False
    
End Sub
    
Public Sub updatedb(post As Integer)
    Dim ext As String
    Dim filnam As String
    Dim strFromFile As String
    Dim lngFileSize As Long
    Dim filenum As Integer
    
    Set Db = OpenDatabase(App.Path & "\klf.mdb")
    Set Rs = Db.OpenRecordset("data")
    On Error Resume Next
    Rs.Move post - 1
    Rs.Edit
    Rs.Fields(0) = Trim(txtZipName.Text)
    Rs.Fields(1) = Trim(txtAuthorName.Text)
    Rs.Fields(4) = Trim(txtComments.Text)
    If picBuffer.Picture > "" Then
        If Seld = True Then
            'rename picture with zip name so they are the same
            ext = Mid$(cmndlg.FileName, InStrRev(cmndlg.FileName, ".", , vbTextCompare) + 1)
            If Len(ext) = 3 Then filnam = LEFT$(txtZipName.Text, Len(txtZipName.Text) - 3)
            If Len(ext) = 4 Then filnam = LEFT$(txtZipName.Text, Len(txtZipName.Text) - 4)
            Rs.Fields(2) = filnam & ext
            
            Call SavePicture(picBuffer.Picture, App.Path & "\tmpfile")
            If PhotosInFolder = 1 Then FileCopy cmndlg.FileName, App.Path & "\Photos\" & filnam & ext    'copy picture into Photo folder
            Seld = False
            KeepCurrentPic = True
            
            filenum = FreeFile
            lngFileSize = FileLen(App.Path & "\tmpfile")
            strFromFile = String(lngFileSize, " ")
            Open App.Path & "\tmpfile" For Binary As filenum
            Get filenum, , strFromFile
            Close filenum
            Rs.Fields(3) = strFromFile
            Kill App.Path & "\tmpfile"
        End If
    End If
    Rs.Update
    Rs.Close: Set Rs = Nothing: Db.Close: Set Db = Nothing
    'disables textboxes but they do not show gray
    Picture3.Enabled = False
    Picture4.Enabled = False
    Picture5.Enabled = False
    
    Open App.Path & "\klf.mdb" For Binary As #1
    g = LOF(1)
    Close #1
    txtRightStatus.Text = "Size of the database : " & Format(g, "###,###,###,##0") & " k"
    visapost post
End Sub
    
Private Sub Frame3_DblClick()
    Frame3.Visible = False    'hide authors list
    txtSearch.Text = ""
End Sub
    
Private Sub List1_Click()
    picBuffer.Picture = LoadPicture()
    Dim oldfindstart As Integer
    If List1.Sorted = False Then
        Set Db = OpenDatabase(App.Path & "\klf.mdb")
        Set Rs = Db.OpenRecordset("data")
        Rs.Close: Set Rs = Nothing: Db.Close: Set Db = Nothing
        visapost List1.ListIndex + 1
        postnr = List1.ListIndex + 1
    Else
        findstart = 0
        oldfindstart = 0
        Call find(Trim(List1.Text), findstart)
    End If
    KillFolder App.Path & "\tempUnzip\", True      'remove any old files
    List2.Clear
    rtb1.Text = ""
    cmdLaunch.ForeColor = &H808080
    cmdLaunch.Enabled = False
End Sub
    
Private Sub List1_DblClick()
    KillFolder App.Path & "\tempUnzip\", True      'remove any old files
    If List1.Text = "" Then Exit Sub
    bzip.Unzip App.Path & "\ZippedFiles\" & List1.Text, App.Path & "\tempUnzip\"
    List2.Clear
    rtb1.Text = ""
    'load list1 with files
    FileList App.Path & "\tempUnzip\"
End Sub
    
Private Sub List2_Click()
    rtb1.Text = ""
    rtb1.Text = FileText(List2.Text)
    ColorIn rtb1                        'color text (syntax)
    cmdLaunch.Enabled = True
    cmdLaunch.ForeColor = vbBlack
End Sub
    
Private Sub List2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lR As Long
    
    lR = (CLng(X / Screen.TwipsPerPixelX) And &HFFFF) Or (&H10000 * CLng(Y / Screen.TwipsPerPixelY))
    lR = SendMessage(List2.hwnd, LB_ITEMFROMPOINT, 0&, ByVal lR)
    If lR > -1 Then
        lR = lR And &H7FFF
        List2.ToolTipText = List2.List(lR)
    End If
End Sub
    
Private Sub lstAuthors_Click()
    Dim lgth As Integer
    
    lgth = Len(lstAuthors.Text)
    If lgth <= 25 Then
        txtSearch.Text = LEFT$(lstAuthors.Text, 3)
    Else
        txtSearch.Text = LEFT$(lstAuthors.Text, 6)
    End If
    LoadAuthorList
End Sub

Private Sub mnuAuthorsList_Click()
   cmdShowAuthors_Click
End Sub

Private Sub mnuDeleteAll_Click()
   cmdDeleteAll_Click
End Sub

Private Sub mnuDeleteFile_Click()
   cmdDeleteCont_Click
End Sub

Private Sub mnuEditFile_Click()
   cmdEditCont_Click
End Sub

Private Sub mnuExit_Click()
   cmdQuit_Click
End Sub

Private Sub mnuKeepPics_Click()
Dim opt As String
  
   If mnuKeepPics.Checked = True Then
      opt = "0"
      mnuKeepPics.Checked = False
      PhotosInFolder = 0
   Else
      opt = "1"
      mnuKeepPics.Checked = True
      PhotosInFolder = 1
   End If
   SaveText opt, App.Path & "\Options.txt"
End Sub

Private Sub mnuNewFile_Click()
   cmdNewCont_Click
End Sub

Private Sub picBuffer_Click()
    KeepCurrentPic = False
    Seld = True
End Sub
    
Private Sub picMain_DblClick()
On Error GoTo SkipMe
    If editing = True Then
        KeepCurrentPic = False
        Seld = True
        ShowOpen
        If cmndlg.FileName = "" Then Exit Sub
        picBuffer.Visible = True
        picMain.Visible = False
        picBuffer.Picture = LoadPicture(cmndlg.FileName)
        txtPicName.Text = Mid$(cmndlg.FileName, InStrRev(cmndlg.FileName, "\") + 1, Len(cmndlg.FileName))  'filename and extension
        If picBuffer.ScaleWidth >= picMain.ScaleWidth Then picBuffer.LEFT = 120
        If picBuffer.ScaleHeight >= picMain.ScaleHeight Then picBuffer.TOP = 240
        If picBuffer.ScaleWidth < picMain.ScaleWidth Then
            picBuffer.LEFT = (picMain.ScaleWidth / 2) - (picBuffer.ScaleWidth / 2)
        End If
        If picBuffer.ScaleHeight < picMain.ScaleHeight Then
            picBuffer.TOP = (picMain.ScaleHeight / 2) - (picBuffer.ScaleHeight / 2)
        End If
        picMain.Cls
    End If
    Seld = True
    KeepCurrentPic = False
    Exit Sub
SkipMe:
    MsgBox "Error. Wrong Format"
End Sub
    
Private Sub picMain_KeyPress(KeyAscii As Integer)
    If editing = True Then
        If KeyAscii Then
            If Clipboard.GetFormat(2) = True Then
                picBuffer.Visible = True
                picMain.Visible = False
                picBuffer.Picture = Clipboard.GetData(2)
                If picBuffer.ScaleWidth > picMain.ScaleWidth Then picBuffer.LEFT = 0
                If picBuffer.ScaleHeight > picMain.ScaleHeight Then picBuffer.TOP = 0
                If picBuffer.ScaleWidth < picMain.ScaleWidth Then
                    picBuffer.LEFT = (picMain.ScaleWidth / 2) - (picBuffer.ScaleWidth / 2)
                End If
                If picBuffer.ScaleHeight < picMain.ScaleHeight Then
                    picBuffer.TOP = (picMain.ScaleHeight / 2) - (picBuffer.ScaleHeight / 2)
                End If
                picMain.Cls
            End If
        End If
    End If
End Sub

Private Sub txtpost_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim post As String
    Set Db = OpenDatabase(App.Path & "\klf.mdb")
    Set Rs = Db.OpenRecordset("data")
    If Rs.RecordCount > 0 Then
        If KeyCode = 13 Then
            If txtpost < 1 Then txtpost = 1
            If txtpost > Rs.RecordCount Then txtpost = Rs.RecordCount
            Rs.Close: Set Rs = Nothing: Db.Close: Set Db = Nothing
            post = txtpost
            visapost txtpost
        End If
    Else
        txtpost = "0"
    End If
End Sub
    
Private Sub txtZipName_KeyPress(KeyAscii As Integer)
    Dim resp As String
    
    If KeyAscii = 13 Then
        If Dir(App.Path & "\ZippedFiles\" & txtZipName.Text) <> "" Then ' it exists already
        resp = MsgBox("Zip filename already exists. Do you want to replace it?", vbYesNo, "Zip FileName already exists")
        If resp = 7 Then      'no
        txtZipName.Text = ""
        txtAuthorName.Text = ""
        cmdNewCont_Click
        Exit Sub
    End If
End If
DeleteFile App.Path & "\ZippedFiles\" & txtZipName.Text            'delete old zip in folder
FileCopy afilepath, App.Path & "\ZippedFiles\" & txtZipName.Text   'copy new zip file into Folder

MsgBox "File saved in ZippedFiles Folder"
End If
End Sub
   
Private Sub txtSearch_Change()
    Set Db = OpenDatabase(App.Path & "\klf.mdb")
    
    If txtSearch.Text = vbNullString Then
        Set Rs = Db.OpenRecordset("data", dbOpenTable)
    Else
        Set Rs = Db.OpenRecordset("SELECT * FROM data WHERE AuthorName LIKE '" & txtSearch.Text & "'" & "& '*'")
    End If
    LoadAuthorList
End Sub
    
Private Function FileList(ByVal Pathname As String, Optional DirCount As Long, Optional FileCount As Long) As String
    Dim ShortName As String, LongName As String
    Dim NextDir As String
    Dim fnExt As String
    On Error Resume Next
    
    Static FolderList As Collection
    Set FolderList = Nothing                     'clear for next file
    Screen.MousePointer = vbHourglass
    If FolderList Is Nothing Then
        Set FolderList = New Collection
        FolderList.add Pathname
        DirCount = 0
        FileCount = 0
    End If
    
    Do
        NextDir = FolderList.Item(1)
        FolderList.Remove 1
        ShortName = Dir(NextDir & "\*.*", vbNormal Or vbArchive Or vbDirectory)
        
        Do While ShortName > ""
            
            If ShortName = "." Or ShortName = ".." Then
            Else
                LongName = NextDir & "\" & ShortName
                fnExt = LCase(RIGHT$(LongName, 4))               'get extension and make lower case
                'skip unnessary files
                If fnExt = ".jpg" Or fnExt = ".jpeg" _
                Or fnExt = ".gif" Or fnExt = ".bmp" _
                Or fnExt = ".ico" Or fnExt = ".frx" _
                Or fnExt = ".scc" Or fnExt = ".ctx" _
                Then GoTo here
                'put files in listbox
                List2.AddItem LongName
here:
                If (GetAttr(LongName) And vbDirectory) > 0 Then
                    FolderList.add LongName
                    DirCount = DirCount + 1
                Else
                    FileList = FileList & LongName & vbCrLf
                    FileCount = FileCount + 1
                End If
            End If
            ShortName = Dir()
        Loop
    Loop Until FolderList.Count = 0
    Screen.MousePointer = vbNormal
End Function

Private Function FileText(ByVal FileName As String) As String
    Dim handle As Integer
    On Error Resume Next
    
    If Len(Dir$(FileName)) = 0 Then
        Err.Raise 53
    End If
    
    handle = FreeFile
    Open FileName$ For Binary As #handle
    FileText = Space$(LOF(handle))
    Get #handle, , FileText
    Close #handle
End Function

Private Sub LoadListbox()
    Dim X As Integer
    List1.Clear
    Set Db = OpenDatabase(App.Path & "\klf.mdb")
    Set Rs = Db.OpenRecordset("data")
    
    For X = 0 To Rs.RecordCount - 1
        List1.AddItem Rs.Fields(0)
        Rs.MoveNext
    Next X
    
    Rs.Close: Set Rs = Nothing: Db.Close: Set Db = Nothing
End Sub

Private Function LoadAuthors()
    Dim errormsg As String
    Dim Max As Integer
    Dim i As Integer
    
    Set Db = OpenDatabase(App.Path & "\klf.mdb")
    Set Rs = Db.OpenRecordset("data", dbOpenTable)
    If Rs.RecordCount = 0 Then
        errormsg = MsgBox("No Records Found", , "Error")
        txtSearch.Text = ""
        Exit Function
    End If
    
    Rs.MoveLast
    Rs.MoveFirst
    Max = Rs.RecordCount
    Rs.MoveFirst
    
    lstAuthors.Clear
    
    For i = 1 To Max
        If txtSearch = "" Then
            lstAuthors.AddItem Rs("AuthorName") & vbTab & " Record Number = " & i
            Rs.MoveNext
        Else
            lstAuthors.AddItem Rs("AuthorName") & vbTab & " Author Count #" & i
            Rs.MoveNext
        End If
    Next i
    Rs.Close: Set Rs = Nothing: Db.Close: Set Db = Nothing
End Function

Private Function LoadAuthorList()
    Dim i As Integer
    Dim lgth As Integer
    
    If txtSearch.Text = "" Then
        lstAuthor2.Clear
        Exit Function
    End If
    
    lstAuthor2.Clear
    lgth = Len(txtSearch.Text)
    For i = 0 To lstAuthors.ListCount - 1
        If LCase(LEFT$(lstAuthors.List(i), lgth)) = LCase(LEFT$(txtSearch.Text, lgth)) Then
            lstAuthor2.AddItem lstAuthors.List(i)
        End If
    Next i
    If lstAuthor2.ListCount = 0 Then lstAuthor2.AddItem "NO MATCH"
End Function

Private Function Shell(Program As String, Optional ShowCmd As Long = vbNormalNoFocus, Optional ByVal WorkDir As Variant) As Long
    Dim FirstSpace As Integer, Slash As Integer
    'this code is by EM Dixson
    On Error Resume Next
    If LEFT(Program, 1) = """" Then
        FirstSpace = InStr(2, Program, """")
        If FirstSpace <> 0 Then
            Program = Mid(Program, 2, FirstSpace - 2) & Mid(Program, FirstSpace + 1)
            FirstSpace = FirstSpace - 1
        End If
    Else
        FirstSpace = InStr(Program, " ")
    End If
    If FirstSpace = 0 Then FirstSpace = Len(Program) + 1
    
    If IsMissing(WorkDir) Then
        For Slash = FirstSpace - 1 To 1 Step -1
            If Mid(Program, Slash, 1) = "\" Then Exit For
        Next
        
        If Slash = 0 Then
            WorkDir = CurDir
        ElseIf Slash = 1 Or Mid(Program, Slash - 1, 1) = ":" Then
            WorkDir = LEFT(Program, Slash)
        Else
            WorkDir = LEFT(Program, Slash - 1)
        End If
    End If
    Shell = ShellExecute(0, vbNullString, _
    LEFT(Program, FirstSpace - 1), LTrim(Mid(Program, FirstSpace)), _
    WorkDir, ShowCmd)
End Function

