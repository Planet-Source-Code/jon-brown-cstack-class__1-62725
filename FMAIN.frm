VERSION 5.00
Begin VB.Form FMAIN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stack Class Example"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   6495
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEmpty 
      Caption         =   "IsEmpty"
      Height          =   350
      Left            =   240
      TabIndex        =   9
      ToolTipText     =   "Is the Stack empty?"
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdPeek 
      Caption         =   "Peek"
      Enabled         =   0   'False
      Height          =   350
      Left            =   240
      TabIndex        =   8
      ToolTipText     =   "View the top item on the Stack"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   350
      Left            =   1080
      TabIndex        =   6
      ToolTipText     =   "Close the program"
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Enabled         =   0   'False
      Height          =   350
      Left            =   1800
      TabIndex        =   5
      ToolTipText     =   "Empty the Stack"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy"
      Height          =   350
      Left            =   1800
      TabIndex        =   4
      ToolTipText     =   "Display the contents of the Stack"
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdPop 
      Caption         =   "Pop"
      Enabled         =   0   'False
      Height          =   350
      Left            =   240
      TabIndex        =   3
      ToolTipText     =   "Pop the top value off the Stack"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdPush 
      Caption         =   "Push"
      Default         =   -1  'True
      Height          =   350
      Left            =   240
      TabIndex        =   2
      ToolTipText     =   "Push a value onto the Stack"
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtDisplay 
      Height          =   2535
      Left            =   3240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   360
      Width           =   3015
   End
   Begin VB.TextBox txtValue 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Enter PUSH value here"
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label lblCount 
      Caption         =   "STACK COUNT: 0"
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "FMAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'//********************************************************************************\\
'  ********************************************************************************
'  FORM:            FMAIN.FRM

'  PURPOSE:         TO DEMONSTRATE THE CSTACK.CLS CLASS MODULE

'  AUTHOR:          Jon S. Brown
'  DATE:            Sept. 30, 2005
'  UPDATED:         Oct. 3, 2005
'  ********************************************************************************
'\\********************************************************************************//

Option Explicit

'// declare a Stack
Dim Stack As New CSTACK

Private Sub Form_Load()
'// set the Stack
    Set Stack = New CSTACK
'// set the input textbox
    txtValue.SelStart = 0
    txtValue.SelLength = Len(txtValue.Text)
End Sub

Private Sub cmdPush_Click()
'// test for a value and assign NULL if empty
    If txtValue.Text = "" Then txtValue.Text = "NULL"
'// push the value onto the stack
    Stack.Push (txtValue.Text)
'// update txtDisplay
    txtDisplay.Text = txtDisplay.Text & "Push: " & Stack.Peek & vbCrLf
'// reset the input textbox
    Reset (False)
'// enable the Pop, Peek, and Clear buttons
    cmdPop.Enabled = True
    cmdPeek.Enabled = True
    cmdClear.Enabled = True
End Sub

Private Sub cmdPop_Click()
'// pop the top item off the Stack and onto txtDisplay
    txtDisplay.Text = txtDisplay.Text & "Pop: " & Stack.Pop & vbCrLf
'// disable buttons if Stack is empty
    If Stack.IsEmpty Then
        cmdPop.Enabled = False
        cmdPeek.Enabled = False
        cmdClear.Enabled = False
    End If
'// reset the input textbox
    Reset (False)
'// enable the Push button
    cmdPush.Enabled = True
End Sub

Private Sub cmdPeek_Click()
'// view the top item on the Stack
    txtDisplay.Text = txtDisplay.Text & "Peek: " & Stack.Peek & vbCrLf
'// reset the input textbox
    Reset (False)
End Sub

Private Sub cmdCopy_Click()
'// an array to hold the stack values
    Dim tempArray() As Variant
'// copy the array
    tempArray = Stack.Copy
'// a counter
    Dim i As Integer
    i = Stack.Count
'// display the contents in txtDisplay
    With txtDisplay
        .Text = Empty
        .Text = "Stack contents:" & vbCrLf
            Do Until i = 0
                .Text = .Text & (i - 1) & ": " & tempArray(i - 1) & vbCrLf
                i = (i - 1)
            Loop
        .Text = .Text & "Stack count = " & Stack.Count & vbCrLf & vbCrLf
    End With
'// reset the input textbox
    Reset (False)
End Sub

Private Sub cmdClear_Click()
'// clear the Stack
    Stack.Clear
'// enable the Push button
    cmdPush.Enabled = True
'// disable the Pop, Peek and Clear buttons
    cmdPop.Enabled = False
    cmdPeek.Enabled = False
    cmdClear.Enabled = False
'// Reset All textboxes
    Reset (True)
End Sub

Private Sub cmdEmpty_Click()
'// test if the Stack is empty
    Select Case Stack.IsEmpty
        Case True
            MsgBox "The Stack is empty!", vbInformation, "Stack Test"
        Case False
            MsgBox "The Stack is not empty!", vbInformation, "Stack Test"
    End Select
'// reset the input textbox
    Reset (False)
End Sub

Private Sub cmdExit_Click()
'// exit the program
    Unload Me
End Sub

Private Sub Reset(Display As Boolean)
'// set focus on input textbox
    txtValue.SetFocus
'// reset input textbox, ready to accept values
    With txtValue
        .Text = "Enter PUSH value here"
        .SelStart = 0
        .SelLength = Len(txtValue.Text)
    End With
'// update the count monitor
    lblCount.Caption = "STACK COUNT: " & Stack.Count

'// if we are displaying the Stack clear txtDisplay first
    If Display Then txtDisplay.Text = Empty
End Sub


