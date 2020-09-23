VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Set QuickLinks"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7665
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   7665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command9 
      Caption         =   "Help"
      Height          =   375
      Left            =   4560
      TabIndex        =   25
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Set"
      Height          =   375
      Left            =   6600
      TabIndex        =   24
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   23
      Text            =   "C:\"
      Top             =   3480
      Width           =   4695
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Set"
      Height          =   375
      Left            =   6600
      TabIndex        =   21
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "C:\"
      Top             =   3000
      Width           =   4695
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Set"
      Height          =   375
      Left            =   6600
      TabIndex        =   18
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "C:\"
      Top             =   2520
      Width           =   4695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Set"
      Height          =   375
      Left            =   6600
      TabIndex        =   15
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "C:\"
      Top             =   2040
      Width           =   4695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Set"
      Height          =   375
      Left            =   6600
      TabIndex        =   12
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "C:\"
      Top             =   1560
      Width           =   4695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Set"
      Height          =   375
      Left            =   6600
      TabIndex        =   9
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "C:\"
      Top             =   1080
      Width           =   4695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Set"
      Height          =   375
      Left            =   6600
      TabIndex        =   6
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "C:\"
      Top             =   600
      Width           =   4695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Set"
      Height          =   375
      Left            =   6600
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "C:\"
      Top             =   120
      Width           =   4695
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Exit"
      Height          =   375
      Left            =   6120
      TabIndex        =   0
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "QuickLink #8:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   3480
      Width           =   2295
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "QuickLink #7:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "QuickLink #6:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "QuickLink #5:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "QuickLink #4:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "QuickLink #3:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "QuickLink #2:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "QuickLink #1:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' The following functions for the 8 Command buttons set the text box on the same
' line with the commmand button so that they show the path of the Directory box
' on Form1.
Private Sub Command1_Click()
Text1.Text = Form1.Dir1.path
End Sub
Private Sub Command2_Click()
Text2.Text = Form1.Dir1.path
End Sub
Private Sub Command3_Click()
Text3.Text = Form1.Dir1.path
End Sub
Private Sub Command4_Click()
Text4.Text = Form1.Dir1.path
End Sub
Private Sub Command5_Click()
Text5.Text = Form1.Dir1.path
End Sub
Private Sub Command6_Click()
Text6.Text = Form1.Dir1.path
End Sub
Private Sub Command7_Click()
Text7.Text = Form1.Dir1.path
End Sub
Private Sub Command8_Click()
Text8.Text = Form1.Dir1.path
End Sub
Private Sub Command10_Click()
' Hides form for later use
Form4.Hide
' Resets the QuickLink menu so it now matches the new text.
Form1.QuickLink1MNU.Caption = Text1.Text
Form1.QuickLink2MNU.Caption = Text2.Text
Form1.QuickLink3MNU.Caption = Text3.Text
Form1.QuickLink4MNU.Caption = Text4.Text
Form1.QuickLink5MNU.Caption = Text5.Text
Form1.QuickLink6MNU.Caption = Text6.Text
Form1.QuickLink7MNU.Caption = Text7.Text
Form1.QuickLink8MNU.Caption = Text8.Text
End Sub

Private Sub Command9_Click()
' Messagebox that shows directions on how to set QuickLinks
MsgBox "To set a QuickLink simply have the desired drive selected on the main form and then press set. The QuickLink should be now set and accessable from the main form's drop down menu.", vbInformation, "QuickLink Help"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Text1.SaveFile App.path, "\mem\1.txt"
Text2.SaveFile App.path, "\mem\2.txt"
Text3.SaveFile App.path, "\mem\3.txt"
Text4.SaveFile App.path, "\mem\4.txt"
Text5.SaveFile App.path, "\mem\5.txt"
Text6.SaveFile App.path, "\mem\6.txt"
Text7.SaveFile App.path, "\mem\7.txt"
Text8.SaveFile App.path, "\mem\8.txt"
End Sub
