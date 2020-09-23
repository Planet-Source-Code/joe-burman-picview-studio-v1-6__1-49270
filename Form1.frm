VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00404040&
   Caption         =   "PicView Studio v1.6 - An Origional jBurman.com Application"
   ClientHeight    =   6690
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10740
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6690
   ScaleWidth      =   10740
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   120
      Width           =   10575
   End
   Begin VB.VScrollBar VScroll1 
      Enabled         =   0   'False
      Height          =   5895
      Left            =   10440
      Min             =   1
      TabIndex        =   5
      Top             =   480
      Value           =   1
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Enabled         =   0   'False
      Height          =   255
      Left            =   3120
      TabIndex        =   4
      Top             =   6360
      Width           =   7335
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   5895
      Left            =   3120
      ScaleHeight     =   5835
      ScaleWidth      =   7275
      TabIndex        =   3
      Top             =   480
      Width           =   7335
      Begin VB.Image Image1 
         Height          =   495
         Left            =   240
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2895
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2340
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   2895
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   3210
      Left            =   120
      Pattern         =   "*.bmp;*.gif;*.art;*.jpg;*.jpeg;*.cur;*.ico;*.icon"
      TabIndex        =   2
      Top             =   3240
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   -720
      TabIndex        =   7
      Top             =   6600
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   -600
      TabIndex        =   6
      Top             =   6480
      Width           =   615
   End
   Begin VB.Menu FileMNU 
      Caption         =   "File"
      Begin VB.Menu OptionsMNU 
         Caption         =   "Options"
      End
      Begin VB.Menu Space3 
         Caption         =   "-"
      End
      Begin VB.Menu ExitMNU 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu ViewMNU 
      Caption         =   "View"
      Begin VB.Menu FullSizeMNU 
         Caption         =   "Full Size"
      End
      Begin VB.Menu StretchMNU 
         Caption         =   "Stretch to Window"
      End
      Begin VB.Menu FitWndMNU 
         Caption         =   "Fit to Window (Proportional)"
      End
      Begin VB.Menu FullScreenMNU 
         Caption         =   "Full Screen"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Drive1_Change()
On Error GoTo 13
' Change the directory list to show changed drive's contents
Dir1.Path = Drive1.Drive
' Display current location on hard drive in the text box.
Text1.Text = Dir1.Path
GoTo 14 'Skips the error message
13 MsgBox "There was an error while attempting to open file. Please reinsert disk and try again.", vbCritical, "Error Opening File."
Drive1.Drive = "C:\"
14 End Sub
Private Sub Dir1_Change()
' If path was opened by QuickLink then change Drive1's Drive setting
Drive1.Drive = Dir1.Path
' Change the file list to show changed directory's contents
File1.Path = Dir1.Path
' Display current location on hard drive in the text box.
Text1.Text = Dir1.Path
End Sub

Private Sub ExitMNU_Click()
Unload Me
End Sub

Private Sub File1_Click()
' This will send the user an error message if there is any type of problem
On Error GoTo 13

' Display the file's location on hard drive in the text box by displaying
' the current directory, adding a backslash, and then displaying the file's name.
Text1.Text = Dir1.Path & "\" & File1.FileName
Image1.Picture = LoadPicture(Dir1.Path & "\" & File1.FileName)


' Reset all the file viewing operations for the new image. Then preform the full
' Image setting by default
Image1.Stretch = False
VScroll1.Enabled = False
HScroll1.Enabled = False
FitWndMNU_Click
GoTo 14 ' This will skip the default error message

13 MsgBox "There was an error while attempting to open file. Please reinsert disk and try again.", vbCritical, "Error Opening File."
14 End Sub

Private Sub FitWndMNU_Click()
' This is the portion of the program that uses the most math functions. (Sigh)
' Don't worry, there are no more scroll bars or anything to deal with here,
' just hard-core math. Hope your ready to dust off your old Algebra text-
' book, because it would be one hell of a help here.

' Here we are just setting all of the image's qualities back to normal, and
' setting the image in the top left corner.
Image1.Stretch = False
Image1.Left = 0
Image1.Top = 0
Image1.Visible = False

' Now, we have to deactivate the scroll bars.
VScroll1.Enabled = False
HScroll1.Enabled = False

' Well okay, here is the first thing we have to figure out. Is the image's width,
' or it's height, the longer of the two? If the width is longer we set
' "Label1.caption" to "1", but if height is longer, "Label1.caption" will become
' "0". This may seem pointless, but later we are going to center the image in it's
' container and wether the image is enlongated upwards or outwards will be of the
' most importance.
    'This conditional is used if the image is more wide than tall.
    If (Image1.Height) < (Image1.Width) Then
        Label1.Caption = "1"
    End If
    'This conditional is used if the image is more tall than wide.
    If (Image1.Height) > (Image1.Width) Then
        Label1.Caption = "0"
    End If

' Now, of course, we are going to have cases where the image's width will equal
' it's length (the image will be a perfect square.) If this is the case, then
' "Label1.caption" will become "2".
    'This conditional is used if the image is as tall as it is wide.
    If (Image1.Height) = (Image1.Width) Then
        Label1.Caption = "2"
    End If

' Woohoo. Great now the fun math stuff. Here we are going to be using fractions
' and percentages. In the words of Chris Farley; "Sweet mother of god...what is
' the hold up?!?!" See, even programmers can make a joke here and there. ;)
' Speaking of human qualities, I have been programming this application for
' almost five hours straight...My eyes hurt, my head is pounding as I down about
' 5 Tylenols, and I'm on my sixth Mountain Dew. Good times. Alright, back to
' business. Haha!

' We start off by finding how much taller, or shorter, the image is (compared to
' it's width.) Were going to use the "Label1.Caption" to determine which function
' we are going to use. After we divide the width by the height, or divide the height
' by the width, the percentage is stored in "Label2.Caption". We can pretty much
' skip this step if our image is a square.
    
    'This is the function used if we found the width larger than the height.
    If Label1.Caption = "1" Then
        Label2.Caption = (Image1.Height) / (Image1.Width)
    End If
        
    'This is the function used if we found the height larger than the width.
    If Label1.Caption = "0" Then
        Label2.Caption = (Image1.Width) / (Image1.Height)
    End If

' Now this part is pretty cool. Again, we use the Label1.Caption. This
' time, we use it to take the reference of if the height or width is
' larger and then we make that measurement equal to the same dimention of the box in
' which the image is contained. Now, here's where those percentages come in handy.
' See, after the image is fitted in to the box, we multiply the latter of the two
' dimentions to restore it to it's origional dimentions. Don't understand? Just watch.

' We have to allow the image's stretch function to change it's shape (duh.)
Image1.Stretch = True

    'The width is larger here.
    If Label1.Caption = "1" Then
        'We make the wider image's width equal the box's width.
        Image1.Width = Picture1.Width
        'Now multiply the height with the origional width comparrison percentage.
        Image1.Height = (Label2.Caption) * (Image1.Width)
        Image1.Top = (Picture1.Height \ 2) - (Image1.Height \ 2)
    End If
    
    'The height is larger here.
    If Label1.Caption = "0" Then
        'We make the taller image's height equal the box's height.
        Image1.Height = Picture1.Height
        'Now multiply the width with the origional height comparrison percentage.
        Image1.Width = (Label2.Caption) * (Image1.Height)
        Image1.Left = (Picture1.Width \ 2) - (Image1.Width \ 2)
    End If
    
    'Now for the equal height and width set.
    If Label1.Caption = "2" Then
        'We make the equal image's height equal the box's height (because the display
        'area is actually more wide than tall (on default, that is).
        Image1.Height = Picture1.Height
        Image1.Width = Image1.Height
        Image1.Left = (Picture1.Width \ 2) - (Image1.Width \ 2)
    End If
Image1.Visible = True
End Sub

Private Sub Form_Load()
' Display current location on hard drive in the text box. (On the program start
' it will display the directory in which the executable is located)
Text1.Text = Dir1.Path
End Sub


Private Sub Form_Resize()
' This is the form resize function, what it will do is put the lists, boxes,
' images, horizontal bars, and what-have-you in proportion so the program can
' be resized to the users choosing without the form looking all nasty.
' ------------------------------------------------------------------------
If WindowState = 1 Then
GoTo 11
Else
Text1.Left = 120
Text1.Top = 120
Text1.Height = 285
Text1.Width = Form1.Width - (240) - 120
Drive1.Top = 480
Drive1.Left = 120
Drive1.Width = 2895
Dir1.Top = 840
Dir1.Left = 120
Dir1.Width = 2895
Dir1.Height = (Form1.Height - 960) \ 2
File1.Left = 120
File1.Width = 2895
File1.Top = (Dir1.Height) + (960)
File1.Height = (Form1.Height) - (Dir1.Height) - (1700)
HScroll1.Width = (Form1.Width) - ((File1.Width) + 735)
HScroll1.Top = (Form1.Height) - 1080
HScroll1.Left = 3120
Picture1.Width = (Form1.Width) - ((File1.Width) + 735)
Picture1.Top = 480
Picture1.Left = 3120
Picture1.Height = (Form1.Height) - 1550
VScroll1.Top = 480
VScroll1.Height = (Form1.Height) - 1550
VScroll1.Left = Picture1.Width + 3120
FitWndMNU_Click
End If
11 End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub FullScreenMNU_Click()
Form2.Show
End Sub

Private Sub FullSizeMNU_Click()
Image1.Visible = False
' Ok guys, this is the hardest function of the program to understand,
' so get ready for some deep thought here, just to understand this.

' Turn off ImageBox's Stretch Feature
Image1.Stretch = False

' Line up the Image's top left corner to the container's top left (zero.)
Image1.Top = 0
Image1.Left = 0

' Enable the vertical scroller if the image is too tall to view in box.
    If (Image1.Height) > (Picture1.Height) Then
        VScroll1.Enabled = True
    End If
    
' Enable the horizontal scroller if the image is too wide to view in box.
    If (Image1.Width) > (Picture1.Width) Then
        HScroll1.Enabled = True
    End If
    
' Now we have to reset the Scroll bars' variables to fit and scroll smoothly
' to the image's width and height.

' Alright, first we have to make it so that the vertical scroll bar can only be
' scrolled up to the top of the image (so it doesn't continue to scroll in to a
' gray abyss.) JUST A HINT: Zero will remain the correct boundary number even if
' the image's container's height is resized.
    If VScroll1.Enabled = True Then
        VScroll1.Min = 0
    End If

' This describes the maximum amount possible to scroll down to. Normally you would
' think all that was needed would be "VScroll1.Max = Image1.Height". But, if I were
' to set the function up like so, it would continue scrolling downwards to a gray,
' blank area, seeing as the the scrollbar will stop scrolling when the picture
' reaches the top of the box. We simply solve this by stopping the scroll bar an
' amount equal to the height of the box. So, now the picture will stop scrolling
' when it hits the bottom of the box.
    If VScroll1.Enabled = True Then
        VScroll1.Max = Image1.Height - Picture1.Height
    End If
    
' Now, we set it up so that the horizontal scoll bar can only be scrolled left
' until the left side of the picture touches the left side of the box, and no
' further. JUST A HINT: Zero will remain the correct boundary number even if
' the image's container's width is resized.
    If HScroll1.Enabled = True Then
        HScroll1.Min = 0
    End If

' Okay, now we set up the image's maximum horizontal scroll amount. This will limit
' the amount of area the picture can move to the right. It is the image's width
' minus the container's width (again, to keep the user from scrolling off in to
' an empty gray space.
    If HScroll1.Enabled = True Then
        HScroll1.Max = Image1.Width - Picture1.Width
    End If

' Geez, okay, this function is hard to explain. What it will do is it will take
' the image's total height, remove the containers height (so you don't continue
' scroll in to a gray empty space) and then divide it by four. So now, every time
' you click the grey bar on the vertical scroll it will change either up, or
' down, by 25 percent (one fourth) of the total picture. Get it?
    If VScroll1.Enabled = True Then
        VScroll1.LargeChange = (Image1.Height - Picture1.Height) \ 4
    End If

' This basically does the exact same thing as the last line of code, only this
' time we are dealing with the horizontal scrollbar, and the width of everything.
' Same deal, clicking on the horizontal bar will change the horizontal scroll either
' left or right by 25 percent.
    If HScroll1.Enabled = True Then
        HScroll1.LargeChange = (Image1.Width - Picture1.Width) \ 4
    End If

' Now, this line of code is quite close to the last two lines. Okay, we are dealing
' with the vertical scrollbar again, but this time we have to edit the amount of
' change in the scrollbar when an arrow button is clicked (rather than the bar itself
' being clicked, which was what we were dealing with in the last two cases.)
' Now, when the up or down arrow is clicked on the vertical scrollbar the pictures
' position is changed by 5 percent (1/20th) of the total picture. This may sound like
' it would take a while to scroll up or down the entire picture, but remember, it
' doesn't just move 5 percent for every click. You can hold down the scroll bar arrow
' in the direction you wish to continously move 5 percent.
    If VScroll1.Enabled = True Then
        VScroll1.SmallChange = (Image1.Height - Picture1.Height) \ 20
    End If

' And now... You guessed it! Now we do the same for the horizontal bar. The arrow
' buttons on the horizontal scroll bar will now move the picture by 5 percent left
' or right.
    If HScroll1.Enabled = True Then
        HScroll1.SmallChange = (Image1.Width - Picture1.Width) \ 20
    End If
    
' Now all we have to do is set the Scroll bars' values at "0". (This just makes sure
' the scroll bar's position on the track coincides with the picture's position in the
' box.
    If HScroll1.Enabled = True Then
        HScroll1.Value = "1"
    End If
    
    If VScroll1.Enabled = True Then
        VScroll1.Value = "1"
    End If
Image1.Visible = True
End Sub

Private Sub HScroll1_Change()
' This just tells the computer that when the scroll bar is moved that the picture
' should be moved to the current value of the scroll bar. The only reason the "* -1"
' is there is because I made a mistake and found out that the scroll bars will move
' opposite to the picture. So, obviously to get an opposite of a number, you have to
' multiply it by negative one. [An example: 5 x (-1) = -5]
Image1.Left = HScroll1.Value * -1
End Sub

Private Sub Label4_Click()
End Sub

Private Sub OptionsMNU_Click()
Form3.Show
End Sub

Private Sub SetQuickMNU_Click()
Form4.Show
End Sub

Private Sub StretchMNU_Click()
Image1.Visible = False
' We have to apply the stretch function to the image.
Image1.Stretch = True
' This will disable the scroll bars seeing as we won't be needing them sense
' the picture will be stretched to the entire container.
VScroll1.Enabled = False
HScroll1.Enabled = False
' Now we stretch the image to the container.
Image1.Top = 0
Image1.Left = 0
Image1.Height = Picture1.Height
Image1.Width = Picture1.Width
Image1.Visible = True
End Sub

Private Sub VScroll1_Change()
' This does the exact same thing as the last line of code, only this time with the
' vertical scrollbar.
Image1.Top = VScroll1.Value * -1
End Sub
