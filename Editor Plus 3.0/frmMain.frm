VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Untitled - Alconsoft - Editor Plus 3.0"
   ClientHeight    =   7215
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   9960
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmMain.frx":0442
   ScaleHeight     =   7215
   ScaleWidth      =   9960
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar Toolbar2 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   9960
      _ExtentX        =   17568
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   16
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SaveAs"
            Object.ToolTipText     =   "Save As"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "Print"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Find"
            Object.ToolTipText     =   "Find"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Undo"
            Object.ToolTipText     =   "Undo"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Date"
            Object.ToolTipText     =   "Date & Time"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   8760
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":074C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0A00
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0B5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0CB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0E0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2B18
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3072
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":35CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":418E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":46E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4C42
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9480
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   6840
      Width           =   9960
      _ExtentX        =   17568
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4657
            MinWidth        =   4657
            Key             =   "status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "2/27/2006"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "9:48 PM"
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   6090
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   720
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   10742
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   3
      TextRTF         =   $"frmMain.frx":4D9C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9360
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4E13
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":536D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":58C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6489
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":69E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6F3D
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7097
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":71F1
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":734B
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":74A5
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":75FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7759
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":78B3
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7A0D
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7B67
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7CC1
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   330
      Width           =   9960
      _ExtentX        =   17568
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bold"
            Object.ToolTipText     =   "Bold"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Italic"
            Object.ToolTipText     =   "Italic"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Underline"
            Object.ToolTipText     =   "Underline"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Strikeout"
            Object.ToolTipText     =   "Strikeout"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "FColor"
            Object.ToolTipText     =   "Fore Colour"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "BColor"
            Object.ToolTipText     =   "Back Colour"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bullet"
            Object.ToolTipText     =   "Bullet"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Left"
            Object.ToolTipText     =   "Left Align"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Center"
            Object.ToolTipText     =   "Center Align"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Right"
            Object.ToolTipText     =   "Right Align"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   3960
         Sorted          =   -1  'True
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   0
         Width           =   2775
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   6840
         TabIndex        =   3
         Text            =   "Combo2"
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open ..."
         Shortcut        =   ^O
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save &As ..."
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuUndo 
         Caption         =   "Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu sp3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "C&ut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "&Select All"
         Shortcut        =   ^A
      End
      Begin VB.Menu sep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFind 
         Caption         =   "&Find"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuFindNext 
         Caption         =   "Find &Next"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuReplace 
         Caption         =   "Replace"
         Shortcut        =   ^H
      End
   End
   Begin VB.Menu mnuFormat 
      Caption         =   "Fo&rmat"
      Begin VB.Menu mnuBold 
         Caption         =   "&Bold"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuItalic 
         Caption         =   "&Italic"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuUnderline 
         Caption         =   "&Underline"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuStrikeout 
         Caption         =   "&Strikeout"
         Shortcut        =   ^W
      End
      Begin VB.Menu sep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBullets 
         Caption         =   "&Bullets"
         Shortcut        =   ^T
      End
      Begin VB.Menu sp10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuParagraph 
         Caption         =   "&Paragraph ..."
         Begin VB.Menu mnuLeftalign 
            Caption         =   "Left Align"
            Shortcut        =   ^L
         End
         Begin VB.Menu mnuCenteralign 
            Caption         =   "Center Align"
            Shortcut        =   ^M
         End
         Begin VB.Menu mnuRightalign 
            Caption         =   "Right Align"
            Shortcut        =   ^R
         End
      End
      Begin VB.Menu sp2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFont 
         Caption         =   "Font ..."
      End
      Begin VB.Menu sep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWordCase 
         Caption         =   "Word Case"
         Begin VB.Menu mnuUpperCase 
            Caption         =   "Upper Case"
            Shortcut        =   ^J
         End
         Begin VB.Menu mnuLowerCase 
            Caption         =   "Lower Case"
            Shortcut        =   ^K
         End
      End
      Begin VB.Menu sp1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuColour 
         Caption         =   "Colour"
         Begin VB.Menu mnuForecolor 
            Caption         =   "&Fore Colour"
            Shortcut        =   {F5}
         End
         Begin VB.Menu mnuBackcolor 
            Caption         =   "&Back Colour"
            Shortcut        =   {F6}
         End
      End
   End
   Begin VB.Menu mnuInsert 
      Caption         =   "&Insert"
      Begin VB.Menu mnuDateTime 
         Caption         =   "Date and Time"
         Shortcut        =   {F4}
      End
      Begin VB.Menu sp22 
         Caption         =   "-"
      End
      Begin VB.Menu mnuObject 
         Caption         =   "Objects ..."
         Shortcut        =   ^Y
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuCalculator 
         Caption         =   "&Calculator"
         Shortcut        =   {F7}
      End
      Begin VB.Menu sep7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPaint 
         Caption         =   "&Paint"
         Shortcut        =   {F8}
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      Begin VB.Menu mnuMinimize 
         Caption         =   "M&inimize"
      End
      Begin VB.Menu mnuMaximize 
         Caption         =   "Ma&ximize"
      End
      Begin VB.Menu mnuRestore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRecentFiles 
         Caption         =   "Recent Files"
         Begin VB.Menu rec1 
            Caption         =   "Untitled"
            Checked         =   -1  'True
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuOnlineHelp 
         Caption         =   "&Online Help"
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu sep8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About Help"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim theFile As String
Dim i, j As Integer
Dim Filter1, Filter2, Filter3, Filter4, myFilter As String
Private TargetPosition As Integer
Dim tempPath As String
Public Findnext As String
Public TargetPos  As Integer
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lparam As Any) As Long

Private Sub Combo1_Click()
     RichTextBox1.SelFontName = Combo1.Text
    RichTextBox1.SelFontSize = Val(Combo2.Text)
End Sub
Private Sub Combo2_Click()
    RichTextBox1.SelFontName = Combo1.Text
    RichTextBox1.SelFontSize = Val(Combo2.Text)
End Sub



Private Sub Form_Load()
    StatusBar1.Panels("status").Text = "Please Wait... (Loading Fonts)"
    StatusBar1.Refresh
    For i = 1 To Screen.FontCount
        If Screen.Fonts(i) <> "" Then Combo1.AddItem Screen.Fonts(i)
    Next i
    For j = 2 To 100 Step 2
        Combo2.AddItem j
    Next j
    Combo1.Text = "Arial"
    Combo2.Text = "8"
    RichTextBox1.Font.Name = Combo1.Text
    RichTextBox1.Font.Size = Combo2.Text
    On Error GoTo tt
    Cmdline = Command()
    tempPath = Cmdline
    RichTextBox1.LoadFile (tempPath)
    If tempPath <> "" Then
        Me.Caption = tempPath & " - Alconsoft - Editor Plus 3.0"
    Else
        Me.Caption = "Untitled - Alconsoft - Editor Plus 3.0"
    End If
    RichTextBox1.DataChanged = False
tt:
    StatusBar1.Panels("status").Text = ""
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    With RichTextBox1
        .Width = Me.Width - 100
        .Height = Me.Height - 1875
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If RichTextBox1.DataChanged Then
        If MsgBox("The text in the untitled file has changed." & vbCrLf & "Do you want to save the changes ?", vbExclamation + vbYesNo, "Alconsoft") = vbYes Then
            mnuSave_Click
        End If
    End If
    End
End Sub



Private Sub mnuAbout_Click()
    frmAbout.Show
End Sub

Private Sub mnuBackcolor_Click()
    On Error Resume Next
    CommonDialog1.Action = 3
    RichTextBox1.BackColor = CommonDialog1.Color
End Sub

Private Sub mnuBold_Click()
    RichTextBox1.SelBold = Not RichTextBox1.SelBold
    RichTextBox1.SetFocus
End Sub

Private Sub mnuBullets_Click()
    With RichTextBox1
        If (IsNull(.SelBullet) = True) Or (.SelBullet = False) Then
            .SelBullet = True
            Toolbar1.Buttons(9).Value = tbrPressed
            mnuBullets.Checked = True
        ElseIf .SelBullet = True Then
            .SelBullet = False
            .SelHangingIndent = False
            Toolbar1.Buttons(9).Value = tbrUnpressed
            mnuBullets.Checked = False
        End If
    End With
End Sub

Private Sub mnuCalculator_Click()
    Shell "calc.exe", vbNormalFocus
End Sub

Private Sub mnuCenteralign_Click()
    Toolbar1.Buttons(11).Value = tbrUnpressed
    Toolbar1.Buttons(12).Value = tbrPressed
    Toolbar1.Buttons(13).Value = tbrUnpressed
    mnuLeftalign.Checked = False
    mnuRightalign.Checked = False
    mnuCenteralign.Checked = True
    RichTextBox1.SelAlignment = rtfCenter
End Sub

Private Sub mnuCopy_Click()
    If RichTextBox1.SelText <> "" Then
    Clipboard.Clear
    Clipboard.SetText RichTextBox1.SelText
    End If
End Sub

Private Sub mnuCut_Click()
    If RichTextBox1.SelText <> "" Then
    Clipboard.Clear
    Clipboard.SetText RichTextBox1.SelText
    RichTextBox1.SelText = ""
    End If
End Sub

Private Sub mnuDateTime_Click()
    RichTextBox1.SelText = Format$(Date, "dd/mm/yyyy") & "" & Time
End Sub

Private Sub mnuDelete_Click()
    RichTextBox1.SelText = ""
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuFind_Click()
    Dialog.Show
End Sub

Private Sub mnuFindNext_Click()
    
    Dialog.Find TargetPos + 1
End Sub

Private Sub mnuFont_Click()
    On Error GoTo tt
    With CommonDialog1
        .Flags = cdlCFBoth
        .Flags = .Flags + cdlCFForceFontExist
        .Flags = .Flags + cdlCFEffects
        .CancelError = True

        
        .ShowFont
        RichTextBox1.SelFontName = .FontName
        RichTextBox1.SelFontSize = .FontSize
        RichTextBox1.SelBold = .FontBold
        RichTextBox1.SelUnderline = .FontUnderline
        RichTextBox1.SelItalic = .FontItalic
        RichTextBox1.SelStrikeThru = .FontStrikethru
        RichTextBox1.SelColor = .Color
        Combo1.Text = .FontName
        Combo2.Text = .FontSize
    End With
tt:
    Exit Sub
       
End Sub

Private Sub mnuForecolor_Click()
    On Error Resume Next
    CommonDialog1.ShowColor
    RichTextBox1.SelColor = CommonDialog1.Color
End Sub

Private Sub mnuItalic_Click()
    RichTextBox1.SelItalic = Not RichTextBox1.SelItalic
End Sub

Private Sub mnuLeftalign_Click()
    Toolbar1.Buttons(11).Value = tbrPressed
    Toolbar1.Buttons(12).Value = tbrUnpressed
    Toolbar1.Buttons(13).Value = tbrUnpressed
    mnuLeftalign.Checked = True
    mnuRightalign.Checked = False
    mnuCenteralign.Checked = False
    RichTextBox1.SelAlignment = rtfLeft
End Sub

Private Sub mnuLowerCase_Click()
    RichTextBox1.SelText = LCase$(RichTextBox1.SelText)
End Sub

Private Sub mnuMaximize_Click()
    Me.WindowState = 2
End Sub

Private Sub mnuMinimize_Click()
    Me.WindowState = 1
End Sub

Private Sub mnuNew_Click()
    On Error Resume Next
    If RichTextBox1.DataChanged Then
        If MsgBox("The text in the untitled file has changed." & vbCrLf & "Do you want to save the changes ?", vbExclamation + vbYesNo, "Alconsoft") = vbYes Then
            mnuSave_Click
        End If
    End If
    tempPath = ""
    RichTextBox1.Text = ""
    RichTextBox1.DataChanged = False
    Me.Caption = "Untitled - Alconsoft - Editor Plus 3.0"
    rec1.Caption = "Untitled"
    CommonDialog1.FileName = ""
End Sub

Private Sub mnuObject_Click()
    CommonDialog1.FileName = ""
    CommonDialog1.DialogTitle = "Open Objects"
    Filter1 = "JPEG Files (*.jpg)|*.jpg"
    Filter2 = "Bitmap Files (*.bmp)|*.bmp"
    Filter3 = "GIF Files (*.gif)|*.gif"
    Filter4 = "All Files (*.*)|*.*"
    myFilter = Filter1 & "|" & Filter2 & "|" & Filter3 & "|" & Filter4
    CommonDialog1.Filter = myFilter
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName <> "" Then
        RichTextBox1.OLEObjects. _
        Add , , CommonDialog1.FileName
    End If
End Sub

Private Sub mnuOnlineHelp_Click()
    Shell "rundll32.exe url.dll,FileProtocolHandler http://geocities.com/alconsoft/editplus.html", 3
End Sub

Private Sub mnuOpen_Click()
    On Error Resume Next
    If RichTextBox1.DataChanged Then
        If MsgBox("The text in the untitled file has changed." & vbCrLf & "Do you want to save the changes ?", vbExclamation + vbYesNo, "Alconsoft") = vbYes Then
            mnuSave_Click
        End If
    End If
        CommonDialog1.FileName = ""
        Filter1 = "Document Files (*.doc)|*.doc"
        Filter2 = "Text Documents (*.txt)|*.txt"
        Filter3 = "Rich Text Format (*.rtf)|*.rtf"
        Filter4 = "All Files (*.*)|*.*"
        myFilter = Filter1 & "|" & Filter2 & "|" & Filter3 & "|" & Filter4
        CommonDialog1.Filter = myFilter
        CommonDialog1.ShowOpen
    If CommonDialog1.FileName <> "" Then
        RichTextBox1.LoadFile (CommonDialog1.FileName)
        tempPath = CommonDialog1.FileName
        Me.Caption = tempPath & " - Alconsoft - Editor Plus 3.0"
        RichTextBox1.DataChanged = False
        rec1.Caption = CommonDialog1.FileName
        RichTextBox1.DataChanged = False
    End If
End Sub



Private Sub mnuPaint_Click()
    Shell "mspaint.exe", vbNormalFocus
End Sub

Private Sub mnuPaste_Click()
    RichTextBox1.SelText = Clipboard.GetText
End Sub

Private Sub mnuPrint_Click()
    On Error GoTo tt
    With CommonDialog1
        .Flags = cdlPDReturnDC + cdlPDNoPageNums
        .CancelError = True
        If RichTextBox1.SelLength = 0 Then
            .Flags = .Flags + cdlPDAllPages
        Else
            .Flags = .Flags + cdlPDSelection
        End If
        .ShowPrinter
        If (.Flags And cdlPDSelection) <> 0 Then
            If RichTextBox1.SelLength <> 0 Then
                RichTextBox1.SelPrint .hDC
            End If
        Else
            RichTextBox1.SelLength = 0
            RichTextBox1.SelPrint .hDC
        End If
    End With
tt:
    Exit Sub
End Sub

Private Sub mnuReplace_Click()
    Dialog.Show
    Dialog.Text1.Text = RichTextBox1.SelText
End Sub

Private Sub mnuRestore_Click()
    Me.WindowState = 0
End Sub

Private Sub mnuRightalign_Click()
    Toolbar1.Buttons(11).Value = tbrUnpressed
    Toolbar1.Buttons(12).Value = tbrUnpressed
    Toolbar1.Buttons(13).Value = tbrPressed
    mnuLeftalign.Checked = False
    mnuRightalign.Checked = True
    mnuCenteralign.Checked = False
    RichTextBox1.SelAlignment = rtfRight
End Sub

Private Sub mnuSave_Click()
    If tempPath = "" Then
        Filter1 = "Document Files (*.doc)|*.doc"
        Filter2 = "Text Documents (*.txt)|*.txt"
        Filter3 = "Rich Text Format (*.rtf)|*.rtf"
        Filter4 = "All Files (*.*)|*.*"
        myFilter = Filter1 & "|" & Filter2 & "|" & Filter3 & "|" & Filter4
        CommonDialog1.Filter = myFilter
        CommonDialog1.ShowSave
        RichTextBox1.SaveFile (CommonDialog1.FileName)
        Me.Caption = CommonDialog1.FileName & " - Alconsoft - Editor Plus 3.0"
        RichTextBox1.DataChanged = False
    Else
        RichTextBox1.SaveFile (tempPath)
    End If
End Sub

Private Sub mnuSaveAs_Click()
    Filter1 = "Document Files (*.doc)|*.doc"
    Filter2 = "Text Documents (*.txt)|*.txt"
    Filter3 = "Rich Text Format (*.rtf)|*.rtf"
    Filter4 = "All Files (*.*)|*.*"
    myFilter = Filter1 & "|" & Filter2 & "|" & Filter3 & "|" & Filter4
    CommonDialog1.Filter = myFilter
    CommonDialog1.ShowSave
    RichTextBox1.SaveFile (CommonDialog1.FileName)
    Me.Caption = CommonDialog1.FileName & " - Alconsoft - Editor Plus 3.0"
    RichTextBox1.DataChanged = False
End Sub

Private Sub mnuSelectAll_Click()
    RichTextBox1.SetFocus
    RichTextBox1.SelStart = 0
    RichTextBox1.SelLength = Len(RichTextBox1.Text)
End Sub

Private Sub mnuStrikeout_Click()
    RichTextBox1.SelStrikeThru = Not RichTextBox1.SelStrikeThru
End Sub

Private Sub mnuUnderline_Click()
    RichTextBox1.SelUnderline = Not RichTextBox1.SelUnderline
End Sub


Private Sub mnuUndo_Click()
    SendKeys "%{BS}"
End Sub


Private Sub mnuUpperCase_Click()
    RichTextBox1.SelText = UCase$(RichTextBox1.SelText)
End Sub

Private Sub RichTextBox1_Change()
    StatusBar1.Panels("status").Text = "Line : " & RichTextBox1.GetLineFromChar(RichTextBox1.SelStart)
End Sub

Private Sub RichTextBox1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RichTextBox1.AutoVerbMenu = True
End Sub

Private Sub RichTextBox1_SelChange()
    RichTextBox1.DataChanged = True
    
    If RichTextBox1.SelAlignment = rtfLeft Then
        Toolbar1.Buttons(11).Value = tbrPressed
        mnuLeftalign.Checked = True
    Else
        Toolbar1.Buttons(11).Value = tbrUnpressed
        mnuLeftalign.Checked = False
    End If
    
    If RichTextBox1.SelAlignment = rtfCenter Then
        Toolbar1.Buttons(12).Value = tbrPressed
        mnuCenteralign.Checked = True
    Else
        Toolbar1.Buttons(12).Value = tbrUnpressed
        mnuCenteralign.Checked = False
    End If
    
    If RichTextBox1.SelAlignment = rtfRight Then
        Toolbar1.Buttons(13).Value = tbrPressed
        mnuRightalign.Checked = True
    Else
        Toolbar1.Buttons(13).Value = tbrUnpressed
        mnuRightalign.Checked = False
    End If
    
    If RichTextBox1.SelBold = True Then
        Toolbar1.Buttons(1).Value = tbrPressed
    Else
        Toolbar1.Buttons(1).Value = tbrUnpressed
    End If
    
    If RichTextBox1.SelItalic = True Then
        Toolbar1.Buttons(2).Value = tbrPressed
    Else
        Toolbar1.Buttons(2).Value = tbrUnpressed
    End If
    
    If RichTextBox1.SelUnderline = True Then
        Toolbar1.Buttons(3).Value = tbrPressed
    Else
        Toolbar1.Buttons(3).Value = tbrUnpressed
    End If
    
    If RichTextBox1.SelStrikeThru = True Then
        Toolbar1.Buttons(4).Value = tbrPressed
    Else
        Toolbar1.Buttons(4).Value = tbrUnpressed
    End If
    
    If RichTextBox1.SelFontName <> "" And RichTextBox1.SelFontSize <> "" Then
        Combo1.Text = RichTextBox1.SelFontName
        Combo2.Text = RichTextBox1.SelFontSize
    Else
        'On Error Resume Next
        Combo1.Text = ""
        Combo2.Text = ""
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Cut"
            mnuCut_Click
        Case "Copy"
            mnuCopy_Click
        Case "Paste"
            mnuPaste_Click
        Case "Delete"
            mnuDelete_Click
        Case "Undo"
            mnuUndo_Click
        Case "Bold"
            mnuBold_Click
        Case "Italic"
            mnuItalic_Click
        Case "Underline"
            mnuUnderline_Click
        Case "Strikeout"
            mnuStrikeout_Click
        Case "FColor"
            mnuForecolor_Click
        Case "BColor"
            mnuBackcolor_Click
        Case "Bullet"
            mnuBullets_Click
        Case "Left"
            mnuLeftalign_Click
        Case "Center"
            mnuCenteralign_Click
        Case "Right"
            mnuRightalign_Click
        Case "Find"
            mnuFind_Click
    End Select
End Sub
Private Sub RichTextBox1_GotFocus()
   On Error Resume Next
   For Each Control In Controls
      Control.TabStop = False
   Next Control
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Cut"
            mnuCut_Click
        Case "Copy"
            mnuCopy_Click
        Case "Paste"
            mnuPaste_Click
        Case "Delete"
            mnuDelete_Click
        Case "Undo"
            mnuUndo_Click
        Case "New"
            mnuNew_Click
        Case "Open"
            mnuOpen_Click
        Case "Save"
            mnuSave_Click
        Case "SaveAs"
            mnuSaveAs_Click
        Case "Date"
            mnuDateTime_Click
        Case "Print"
            mnuPrint_Click
        Case "Find"
            mnuFind_Click
    End Select
End Sub
