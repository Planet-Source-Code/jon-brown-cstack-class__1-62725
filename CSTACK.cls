VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSTACK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'//********************************************************************************\\
'  ********************************************************************************
'  CLASS MODULE:    CSTACK.CLS

'  PURPOSE:         Simulates a Stack environment to hold a data array. The Stack
'                   works under the "Last on - First off" principle.
'  METHODS:
'                   Initialize: Sets initial stack values.

'                   Push:    Takes a Variant parameter and "Pushes" its value onto
'                            the top of the Stack.
'                   Pop:     Returns a Variant containing the top data item off
'                            the Stack.
'                   Peek:    Returns a Variant containing the top data item
'                            without "Popping" it off the stack.
'                   Clear:   Empties the Stack and reinitializes it.
'                   Copy:    Returns a Dynamic Array of Variants containing a copy
'                            of the Stack.
'                   Count:   Returns a Long containing the current number of items
'                            on the Stack.
'                   IsEmpty: Returns a Boolean. True if the Stack is empty.

'  AUTHOR:          Jon S. Brown
'  DATE:            Sept. 11, 2005
'  UPDATED:         Oct. 3, 2005
'  ********************************************************************************
'\\********************************************************************************//

Option Explicit

'// PRIVATE MEMBER VARIABLES
Private Const STACK_INCREMENT = 10 '// sets the increment size of the stack
                                   '   ** DO NOT SET THIS TO LESS THAN ZERO **
Private stack_inc_count As Long '// tracks how many times we've incremented the stack.
Private stack_top As Long       '// tracks how many items are currently on the stack
Private stack_data() As Variant '// dynamic array to hold stack data

Private Sub Class_Initialize()
'// STACK_INCREMENT TEST
' -----------------------------------------------------------------------------------
'// NOTE: THIS TEST WILL WARN YOU TO CHANGE INVALID VALUES, BUT THE THE STACK WILL
'         STILL BE CREATED. IT WILL BE INVALID AND CAUSE ERRORS. CHANGE THE VALUE!
' -----------------------------------------------------------------------------------
    If STACK_INCREMENT <= 0 Then
        StackException (1)
        Exit Sub
    End If
' -----------------------------------------------------------------------------------

'// initial stack holds x items as defined in STACK_INCREMENT
    ReDim stack_data(STACK_INCREMENT)
'// count is 0
    stack_top = 0
'// increment counter is 1
    stack_inc_count = 1
End Sub

Public Sub Push(vNewData As Variant)
'// if no room left in the current increment
    If (stack_top) = (stack_inc_count * STACK_INCREMENT) Then
    '// add another x elements as defined in STACK_INCREMENT
        ReDim Preserve stack_data(stack_top + STACK_INCREMENT)
    '// and add 1 to the increment count
        stack_inc_count = (stack_inc_count + 1)
    End If

'// assign the data
    stack_data(stack_top) = vNewData
'// increment stack_top
    stack_top = (stack_top + 1)
End Sub

Public Function Pop() As Variant
'// check for and handle stack underflow
    If IsEmpty Then
        StackException (0)
        Exit Function
    End If

'// if we can shrink the stack without losing data
    If (stack_top) = ((stack_inc_count - 1) * STACK_INCREMENT) Then
'// do so
        ReDim Preserve stack_data((stack_inc_count * STACK_INCREMENT) - STACK_INCREMENT)
        stack_inc_count = (stack_inc_count - 1)
    End If

'// decrement stack_top by 1
    stack_top = (stack_top - 1)
'// return the value from the top of the stack
    Pop = stack_data(stack_top)
End Function

Public Function Peek() As Variant
'// check that the Stack isn't empty
    If stack_top = 0 Then
        StackException (0)
        Exit Function
    End If
    
'// return top item
    Peek = stack_data(stack_top - 1)
End Function

Public Function Clear()
'// empty the Stack by resetting to intial values
    ReDim stack_data(STACK_INCREMENT)
    stack_top = 0
    stack_inc_count = 1
End Function

Public Function Copy() As Variant()
'// declare and initialize counter
    Dim i As Integer
    i = 0
'// return each data element
    Do While i < stack_top
        Copy = stack_data
        i = (i + 1)
    Loop
End Function

Public Function Count() As Integer
'// return the current count of items
    Count = stack_top
End Function

Private Sub StackException(iCode As Integer)
'// an Stack Exception has been triggered
'// 0 = Empty; 1 = Invalid Increment Value
    Select Case iCode
        Case 0
            MsgBox "Stack is empty.", vbInformation, "Stack Protection"
        Case 1
            MsgBox "Invalid Stack. Set STACK_INCREMENT to a value greater than zero.", vbCritical, "Stack Protection"
    End Select
End Sub

Public Function IsEmpty() As Boolean
'// return True if the Stack is empty
    If stack_top = 0 Then
        IsEmpty = True
    Else
        IsEmpty = False
    End If
End Function

