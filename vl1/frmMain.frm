VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   Caption         =   "[vL] test"
   ClientHeight    =   4695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   ScaleHeight     =   4695
   ScaleWidth      =   5880
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox rtbMain 
      Height          =   3135
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   5530
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmMain.frx":0000
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Load"
      Default         =   -1  'True
      Height          =   495
      Left            =   4920
      TabIndex        =   2
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox txtFilename 
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Text            =   "c:\boot.ini"
      Top             =   360
      Width           =   5535
   End
   Begin VB.Label Label1 
      Caption         =   "Filename:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdLoad_Click()
    'an array that will carry the file with crlf's indicating changing the
    'element
    Dim stTextArray() As String
    'loop variable
    Dim iIndex As Integer
    
    'Handles errors (most likely resulting from incorrect text name)
    On Error GoTo error
    
    'Loads the file into the text box
    rtbMain.LoadFile txtFilename.Text
    
    'Puts the text into a variable
    stTextArray = Split(rtbMain.Text, vbCrLf)
    
    'Clear out the textbox
    rtbMain.Text = ""
    
    'Loop through the array adding the lines and making every second line bold
    For iIndex = LBound(stTextArray) To UBound(stTextArray)
        With rtbMain
            .SelStart = Len(.Text)
            .SelLength = 0
            .SelText = stTextArray(iIndex) & vbCrLf
            If ((iIndex / 2) = Int(iIndex / 2)) Then
                .SelBold = True
            Else
                .SelBold = False
            End If
        End With
    Next iIndex

    txtFilename.SetFocus
    txtFilename.SelStart = 0
    txtFilename.SelLength = Len(txtFilename.Text)
    Focus
    Exit Sub
    
error:
    MsgBox "There was an error, check your path name and try again", vbCritical, "Error!"
    Focus
End Sub

Private Sub Form_Activate()
    Focus
End Sub

Private Sub Focus()
    'Focuses and selects everything in the textbox
    txtFilename.SetFocus
    txtFilename.SelStart = 0
    txtFilename.SelLength = Len(txtFilename.Text)

End Sub

