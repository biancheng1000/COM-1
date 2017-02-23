VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3585
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5220
   LinkTopic       =   "Form1"
   ScaleHeight     =   3585
   ScaleWidth      =   5220
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   2520
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   0
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Sub"
      Height          =   735
      Left            =   2520
      TabIndex        =   3
      Top             =   2520
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   600
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "First Number"
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Second Number"
      Height          =   735
      Left            =   240
      TabIndex        =   6
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Result"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim comAssembly As ComInteropExample.ComInteropClass
Dim i As Integer
Dim a As Integer
Dim b As Integer
Dim c As Integer
'For i = 0 To 20

Set comAssembly = New ComInteropExample.ComInteropClass

a = Val(Text1.Text)
b = Val(Text2.Text)

'Additon

Text3.Text = comAssembly.PerformAddition(a, b)
'Next i
'omAssembly = Nothing

End Sub

Private Sub Command2_Click()

Dim comAssembly As ComInteropExample.ComInteropClass

Set comAssembly = New ComInteropExample.ComInteropClass
Dim a As Integer
Dim b As Integer
Dim c As Integer

a = Val(Text1.Text)
b = Val(Text2.Text)

'Subtraction

c = comAssembly.PerformDeletion(a, b)
Text3.Text = c


End Sub


