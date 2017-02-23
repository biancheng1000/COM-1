VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3735
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3735
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Sub"
      Height          =   735
      Left            =   2640
      TabIndex        =   3
      Top             =   2760
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Result"
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Second Number"
      Height          =   735
      Left            =   360
      TabIndex        =   6
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "First Number"
      Height          =   735
      Left            =   360
      TabIndex        =   5
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim comAssembly As iInterface
Dim i As Integer
Dim a As Integer
Dim b As Integer
Dim c As Integer
'For i = 0 To 20

Set comAssembly = New ComPlusClass


a = Val(Text1.Text)
b = Val(Text2.Text)

'Additon
c = comAssembly.PerformAddition(a, b)
Text3.Text = c
'Next i
'omAssembly = Nothing

End Sub

Private Sub Command2_Click()

Dim comAssembly As iInterface

Set comAssembly = New ComPlusClass
Dim a As Integer
Dim b As Integer
Dim c As Integer

a = Val(Text1.Text)
b = Val(Text2.Text)

'Subtraction

c = comAssembly.PerformSubtraction(a, b)
Text3.Text = c


End Sub

