VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find"
   ClientHeight    =   1005
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5055
   Icon            =   "Dialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1005
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1080
      TabIndex        =   5
      Top             =   600
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Replace"
      Height          =   285
      Left            =   3840
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Find Next"
      Height          =   285
      Left            =   3840
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Replace :"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "&Find What :"
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2

Public TargetPosition As Integer

Private Sub Command1_Click()
    Find TargetPosition + 1
    Form1.Findnext = Text1.Text
End Sub

Private Sub Command3_Click()
    Form1.RichTextBox1.SelText = Text2.Text
End Sub

Private Sub Form_Load()
    On Error Resume Next
    If Form1.RichTextBox1.SelText <> "" Then
        Text1.Text = Form1.RichTextBox1.SelText
    End If
    Me.Top = Form1.Top
    Me.Left = Form1.Left
    SetWindowPos hwnd, _
            HWND_TOPMOST, 0, 0, 0, 0, _
            SWP_NOMOVE + SWP_NOSIZE
End Sub
Public Sub Find(ByVal start As Integer)
Dim pos As Integer
Dim target As String

    target = Text1.Text
    pos = InStr(start, Form1.RichTextBox1.Text, target)
    If pos > 0 Then
        TargetPosition = pos
        Form1.TargetPos = TargetPosition
        Form1.RichTextBox1.SelStart = TargetPosition - 1
        Form1.RichTextBox1.SelLength = Len(target)
        Form1.RichTextBox1.SetFocus
    Else
        MsgBox "Editor Plus has finished searching the document.", vbInformation, "Alconsoft - Note"
        Form1.RichTextBox1.SetFocus
    End If
End Sub


