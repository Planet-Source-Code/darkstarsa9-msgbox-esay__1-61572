VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5100
   LinkTopic       =   "Form1"
   ScaleHeight     =   2775
   ScaleWidth      =   5100
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   240
      Top             =   2400
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   4080
      Top             =   2400
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Exclamation (  ! )"
      Height          =   495
      Left            =   3360
      TabIndex        =   5
      Top             =   1200
      Width           =   1695
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Error"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Information"
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C00000&
      Caption         =   "Test"
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   2280
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   600
      Width           =   4095
   End
   Begin VB.Label Label3 
      Caption         =   "Note :  You must click on a radio buton to make this form work"
      Height          =   495
      Left            =   360
      TabIndex        =   8
      Top             =   1800
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "Caption  :"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Msg  :"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   720
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim information As String
Dim error As String
Dim question As String


If Option1.Value = True Then
information = vbInformation
MsgBox (message), information, (Text2.Text)
End If

If Option2.Value = True Then
error = vbCritical
MsgBox (Text1.Text), error, (Text2.Text)
End If

If Option3.Value = True Then
question = vbExclamation
MsgBox (Text1.Text), question, (Text2.Text)
End If
End Sub

Private Sub Form_Load()
Dim message As String
Dim caption1 As String
message = InputBox("Please enter your message to be displayed")
Text1.Text = message

caption1 = InputBox("Please enter your caption to be displayed")
Text2.Text = caption1

Timer2.Enabled = True
End Sub


Private Sub Timer1_Timer()
Label3.ForeColor = &HC00000
Timer1.Enabled = False
Timer2.Enabled = True
End Sub

Private Sub Timer2_Timer()
Label3.ForeColor = &HC0&
Timer1.Enabled = True
Timer2.Enabled = False
End Sub
