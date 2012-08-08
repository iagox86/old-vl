VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VL Test 2!"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   5355
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstResults 
      Height          =   2010
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   2160
      Width           =   5175
   End
   Begin VB.CommandButton cmdCount 
      Caption         =   "&Count!"
      Default         =   -1  'True
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txt2 
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txt1 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Second number:"
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "First number:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strNumberNames(0 To 9) As String

Private Sub cmdCount_Click()
    'Loop variable:
    Dim iIndex As Integer
    Dim strOutput As String
    Dim strThisNumber As String
    


    If Val(txt1.Text) > Val(txt2.Text) Then
        MsgBox "The first number has to be smaller", vbCritical
        Exit Sub
    End If
    
    If Not (IsNumeric(txt1.Text) And IsNumeric(txt2.Text)) Then
        MsgBox "The boxes must contain numeric values", vbCritical
        Exit Sub
    End If
    
    If Val(txt1.Text) < 0 Or Val(txt2.Text) < 0 Then
        MsgBox "Both numbers must be above zero", vbCritical
        Exit Sub
    End If
    
    If Val(txt1.Text) > 999 Or Val(txt2.Text) > 999 Then
        MsgBox "Both numbers must be less than 1000.", vbCritical
        Exit Sub
    End If
    
    lstResults.Clear
    
    For iIndex = Val(txt1.Text) To Val(txt2.Text)
        strOutput = ""
        strThisNumber = Str(iIndex)
        strThisNumber = Replace(strThisNumber, " ", "")
        While Len(strThisNumber) < 3
            strThisNumber = "0" & strThisNumber
        Wend
        
        'Hundreds are easy:
        If Left(strThisNumber, 1) <> "0" Then
            strOutput = strOutput & strNumberNames(Val(Left(strThisNumber, 1))) & " hundred "
        End If
        
        'Tens I have to be more careful with, mostly because of 10-19
        'I figure it's easier to do this one by hand :(
        Select Case Mid(strThisNumber, 2, 1)
        Case "2"
            strOutput = strOutput & "twenty "
        Case "3"
            strOutput = strOutput & "thirty "
        Case "4"
            strOutput = strOutput & "fourty "
        Case "5"
            strOutput = strOutput & "fifty "
        Case "6"
            strOutput = strOutput & "sixty "
        Case "7"
            strOutput = strOutput & "seventy "
        Case "8"
            strOutput = strOutput & "eighty "
        Case "9"
            strOutput = strOutput & "ninety "
        End Select
        
        'ones are a pain because of "0"
        If Mid(strThisNumber, 2, 1) <> "1" And Mid(strThisNumber, 3, 1) <> "0" Then
            strOutput = strOutput & strNumberNames(Val(Right(strThisNumber, 1)))
        ElseIf Mid(strThisNumber, 2, 1) = "1" Then
            Select Case Right(strThisNumber, 1)
            Case "0"
                strOutput = strOutput & "ten"
            Case "1"
                strOutput = strOutput & "eleven"
            Case "2"
                strOutput = strOutput & "twelve"
            Case "3"
                strOutput = strOutput & "thirteen"
            Case "4"
                strOutput = strOutput & "fourteen"
            Case "5"
                strOutput = strOutput & "fifteen"
            Case "6"
                strOutput = strOutput & "sixteen"
            Case "7"
                strOutput = strOutput & "seventeen"
            Case "8"
                strOutput = strOutput & "eighteen"
            Case "9"
                strOutput = strOutput & "nineteen"
            End Select
        ElseIf iIndex = 0 Then
            strOutput = strOutput & "zero"
        End If
            
        lstResults.AddItem strOutput
    Next
        
End Sub

Private Sub Form_Load()
    strNumberNames(0) = "zero"
    strNumberNames(1) = "one"
    strNumberNames(2) = "two"
    strNumberNames(3) = "three"
    strNumberNames(4) = "four"
    strNumberNames(5) = "five"
    strNumberNames(6) = "six"
    strNumberNames(7) = "seven"
    strNumberNames(8) = "eight"
    strNumberNames(9) = "nine"
    
       
End Sub
