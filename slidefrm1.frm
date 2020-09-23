VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Slideshow viewer"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   425
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   609
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   500
      Left            =   120
      Max             =   3000
      Min             =   500
      SmallChange     =   500
      TabIndex        =   10
      Top             =   1080
      Value           =   500
      Width           =   1335
   End
   Begin VB.CommandButton cmdFullScr 
      Caption         =   "Full Screen"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton CmdNext 
      Caption         =   "Next"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton CmdStop 
      Caption         =   "Stop Slideshow"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Left            =   3840
      Top             =   7440
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start Slideshow"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4200
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   4
      Top             =   7560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.FileListBox File1 
      Height          =   3015
      Left            =   1560
      Pattern         =   "*.bmp;*.jpg;*.jpeg;*.gif"
      TabIndex        =   3
      Top             =   3240
      Width           =   2655
   End
   Begin VB.DirListBox Dir1 
      Height          =   2790
      Left            =   1560
      TabIndex        =   2
      Top             =   480
      Width           =   2655
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   6165
      Left            =   4320
      ScaleHeight     =   411
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   315
      TabIndex        =   0
      Top             =   120
      Width           =   4725
   End
   Begin VB.Label lblHeight 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Height:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label lblWidth 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Width:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Slideshow Speed"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1440
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function StretchBlt Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal hSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Dim SlideSpeed As Long

Private Sub cmdStart_Click()
    SlideSpeed = HScroll1.Value
    Timer1.Interval = SlideSpeed
End Sub

Private Sub cmdStop_Click()
    Timer1.Interval = 0
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo err:
    Dim TempIndex
    TempIndex = File1.ListIndex
    If TempIndex > File1.ListCount Or TempIndex < 0 Then Exit Sub
        If Mid$(File1.Path, Len(File1.Path), 1) = "\" Then
            Kill File1.Path + File1.FileName
        Else
            Kill File1.Path + "\" + File1.FileName
        End If
        File1.Refresh
        If File1.ListCount > TempIndex Then
            File1.ListIndex = TempIndex
        Else
            File1.ListIndex = TempIndex - 1
        End If
    Exit Sub
err:
    TempIndex = File1.ListIndex
        If err.Number = 53 Then
            File1.Refresh
            If File1.ListCount > TempIndex Then
                File1.ListIndex = TempIndex
            Else
                File1.ListIndex = TempIndex - 1
            End If
            Exit Sub
        End If
End Sub

Private Sub cmdNext_Click()
    Dim TempIndex
    TempIndex = File1.ListIndex
    If TempIndex < 0 Then Exit Sub
        If File1.ListCount - 1 = TempIndex Then
            File1.ListIndex = 0
            Exit Sub
        End If
    File1.ListIndex = File1.ListIndex + 1
End Sub

Private Sub cmdFullScr_Click()
    Form2.Width = Screen.Width
    Form2.Height = Screen.Height
    Form2.Picture1.Width = Screen.Width / Screen.TwipsPerPixelX
    Form2.Picture1.Height = Screen.Height / Screen.TwipsPerPixelY
    Form2.Left = 0
    Form2.Top = 0
    Form2.Show
    Form2.Visible = True
    If File1.ListIndex >= 0 Then Call File1_Click
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
        If File1.ListCount <> 0 Then
            File1.ListIndex = 0
        End If
End Sub

Private Sub Drive1_Change()
    On Error GoTo err
    Dir1.Path = Drive1
err:
    If err.Number = 68 Then MsgBox "Drive Not Ready"
End Sub

Private Sub File1_Click()
    Dim DestWidth As Long
    Dim DestHeight As Long
    Dim DestX As Long
    Dim DestY As Long
    Dim retval As Long
    Dim TempIndex
    On Error GoTo err
        If Mid$(File1.Path, Len(File1.Path), 1) = "\" Then
            Picture2.Picture = LoadPicture(File1.Path + File1.FileName)
        Else
            Picture2.Picture = LoadPicture(File1.Path + "\" + File1.FileName)
        End If
    lblWidth.Caption = "Width:" + Str(Picture2.Width)
    lblHeight.Caption = "Height:" + Str(Picture2.Height)
    Picture1.Cls
    Form2.Picture1.Cls
        If Picture2.Width > Picture2.Height Then
            If Form2.Visible = True Then
                DestWidth = Form2.Picture1.Width
                DestHeight = (Form2.Picture1.Width / Picture2.Width) * Picture2.Height
                DestY = (Form2.Picture1.Height - DestHeight) / 2
                retval = StretchBlt(Form2.Picture1.hdc, 0, DestY, DestWidth, DestHeight, Picture2.hdc, 0, 0, Picture2.Width, Picture2.Height, &HCC0020)
            Else
                DestWidth = Picture1.Width
                DestHeight = (Picture1.Width / Picture2.Width) * Picture2.Height
                DestY = (Picture1.Height - DestHeight) / 2
                retval = StretchBlt(Picture1.hdc, 0, DestY, DestWidth, DestHeight, Picture2.hdc, 0, 0, Picture2.Width, Picture2.Height, &HCC0020)
            End If
        Else
            If Form2.Visible = True Then
                DestHeight = Form2.Picture1.Height
                DestWidth = (Form2.Picture1.Height / Picture2.Height) * Picture2.Width
                DestX = (Form2.Picture1.Width - DestWidth) / 2
                retval = StretchBlt(Form2.Picture1.hdc, DestX, 0, DestWidth, DestHeight, Picture2.hdc, 0, 0, Picture2.Width, Picture2.Height, &HCC0020)

            Else
                DestHeight = Picture1.Height
                DestWidth = (Picture1.Height / Picture2.Height) * Picture2.Width
                DestX = (Picture1.Width - DestWidth) / 2
                retval = StretchBlt(Picture1.hdc, DestX, 0, DestWidth, DestHeight, Picture2.hdc, 0, 0, Picture2.Width, Picture2.Height, &HCC0020)
            End If
        End If
    Exit Sub
err:
    TempIndex = File1.ListIndex
        If err.Number = 53 Then
            File1.Refresh
            If File1.ListCount > TempIndex Then
                File1.ListIndex = TempIndex
            Else
                If File1.ListIndex <> -1 Then
                    File1.ListIndex = TempIndex - 1
                End If
            End If
            Exit Sub
        End If
End Sub

Private Sub Form_Load()
    Form1.Left = (Screen.Width - Form1.Width) / 2
    Form1.Top = (Screen.Height - Form1.Height) / 2
End Sub

Private Sub Form_Terminate()
    Timer1.Interval = 0
    Unload Form2
    Unload Form1
    End
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Timer1.Interval = 0
    Unload Form2
    Unload Form1
    End
End Sub

Private Sub HScroll1_Change()
    SlideSpeed = HScroll1.Value
    If Timer1.Interval <> 0 Then Timer1.Interval = SlideSpeed
End Sub

Private Sub Timer1_Timer()
    If File1.ListIndex = -1 Then Exit Sub
        If File1.ListIndex <> File1.ListCount - 1 Then
            File1.ListIndex = File1.ListIndex + 1
        Else
            File1.ListIndex = 0
        End If
    DoEvents
End Sub
