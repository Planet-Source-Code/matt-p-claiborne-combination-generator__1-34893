VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Combination Solver"
   ClientHeight    =   3465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3945
   LinkTopic       =   "Form1"
   ScaleHeight     =   3465
   ScaleWidth      =   3945
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtStart 
      Height          =   285
      Left            =   1080
      TabIndex        =   10
      Top             =   3000
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   3615
      Begin VB.Label Label4 
         Caption         =   "Time To Complete: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label Label3 
         Caption         =   "Completed/Sec: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   3375
      End
      Begin VB.Label Label2 
         Caption         =   "Possible: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label Label1 
         Caption         =   "Completed: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.TextBox txtSave 
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Text            =   "c:\blah.txt"
      Top             =   480
      Width           =   2295
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Log to File..."
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3600
      Top             =   2160
   End
   Begin VB.TextBox txtLen 
      Height          =   285
      Left            =   2880
      TabIndex        =   2
      Text            =   "4"
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Do it."
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   3495
   End
   Begin VB.TextBox txtString 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "abcdefghijklmnopqrstuvwxyz012345679ABCDEFGHIJKLMNOPQRSTUVWXYZ"
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label5 
      Caption         =   "Start from:"
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   3000
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Total As Double
Dim Completed As Double
Dim LastCompleted As Long
Dim LogToFile As Boolean
Dim PrintOutput As Boolean
Dim STRChars As String
Dim MapChars() As String * 1
Dim LenSTRChars As Long
Private Sub Command1_Click()
Dim X As Integer
Dim StartTime As Long
Command1.Enabled = False
Total = 0
Completed = 0
'txtOutput.Text = ""
Timer1.Enabled = True

If Check1.Value Then
    Open txtSave.Text For Output As #1
    LogToFile = True
Else
    LogToFile = False
End If

Check1.Enabled = False


StartTime = GetTickCount

For X = 0 To txtLen - 1
    Total = Total + (Len(txtString) ^ (txtLen - X))
Next
Label1.Caption = "Completed: " & 0
Label2.Caption = "Possible: " & Total
Label3.Caption = "Completed/Sec: " & 0
Label4.Caption = "Time To Complete: " & 0 & " Seconds"

STRChars = txtString.Text
ReDim MapChars(Len(STRChars)) As String * 1
For X = 1 To Len(STRChars)
    MapChars(X) = Mid(STRChars, X, 1)
Next
LenSTRChars = Len(STRChars)
Combinations txtStart.Text, txtLen


Timer1.Enabled = False
Label1.Caption = "Completed: " & Completed
Label4.Caption = "Time To Complete: " & 0 & " Seconds"

If Check1.Value Then
    Close 1
End If

Check1.Enabled = True
'Check2.Enabled = True
Command1.Enabled = True
End Sub

Private Function Combinations(ByVal X As String, ByVal Y As Byte)
Dim P As Byte
Dim STRblah As String


For P = 1 To LenSTRChars
STRblah = X & MapChars(P)


If LogToFile Then
    Print #1, STRblah
End If


If Not Len(STRblah) = Y Then
    Combinations STRblah, txtLen
End If

Completed = Completed + 1
Next



DoEvents
End Function



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Me
End
End Sub

Private Sub Timer1_Timer()
Dim CompSec As Long

CompSec = (Completed - LastCompleted) * (1000 / Timer1.Interval)

Label1.Caption = "Completed: " & Completed
Label2.Caption = "Possible: " & Total
Label3.Caption = "Completed/Sec: " & CompSec
Label4.Caption = "Time To Complete: " & Int((Total - Completed) / CompSec) & " Seconds"
LastCompleted = Completed
End Sub
