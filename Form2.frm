VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   7380
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10125
   LinkTopic       =   "Form2"
   MouseIcon       =   "Form2.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   7380
   ScaleWidth      =   10125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   375
      Left            =   9000
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   8880
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   2295
      Left            =   120
      MouseIcon       =   "Form2.frx":030A
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
Form2.WindowState = 2
' Load the image that is selected on Form1
Image1.Picture = LoadPicture(Form1.Dir1.Path & "\" & Form1.File1.FileName)
GoTo 14
14 End Sub

Private Sub Form_Resize()


' Here we are just setting all of the image's qualities back to normal, and
' setting the image in the top left corner.
Image1.Stretch = False
Image1.Left = 0
Image1.Top = 0


    'This conditional is used if the image is more wide than tall.
    If (Image1.Height) < (Image1.Width) Then
        Label1.Caption = "1"
    End If
    'This conditional is used if the image is more tall than wide.
    If (Image1.Height) > (Image1.Width) Then
        Label1.Caption = "0"
    End If


    'This conditional is used if the image is as tall as it is wide.
    If (Image1.Height) = (Image1.Width) Then
        Label1.Caption = "2"
    End If


    
    'This is the function used if we found the width larger than the height.
    If Label1.Caption = "1" Then
        Label2.Caption = (Image1.Height) / (Image1.Width)
    End If
        
    'This is the function used if we found the height larger than the width.
    If Label1.Caption = "0" Then
        Label2.Caption = (Image1.Width) / (Image1.Height)
    End If


Image1.Stretch = True

    'The width is larger here.
    If Label1.Caption = "1" Then
        'We make the wider image's width equal the form's width.
        Image1.Width = Form2.Width
        'Now multiply the height with the origional width comparrison percentage.
        Image1.Height = (Label2.Caption) * (Image1.Width)
        Image1.Top = (Form2.Height \ 2) - (Image1.Height \ 2)
    End If
    
    'The height is larger here.
    If Label1.Caption = "0" Then
        'We make the taller image's height equal the form's height.
        Image1.Height = Form2.Height
        'Now multiply the width with the origional height comparrison percentage.
        Image1.Width = (Label2.Caption) * (Image1.Height)
    Image1.Left = (Form2.Width \ 2) - (Image1.Width \ 2)
    End If
    
    'Now for the equal height and width set.
    If Label1.Caption = "2" Then
        'We make the equal image's height equal the form's height (because the display
        'area of a monitor is actually more wide than tall (on default, that is).
        Image1.Height = Form2.Height
        Image1.Width = Image1.Height
        Image1.Left = (Form2.Width \ 2) - (Image1.Width \ 2)
    End If
End Sub

Private Sub Image1_Click()
Unload Me
End Sub
