VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " 32 bpp BMP Demo (XP-Alpha.bmp)"
   ClientHeight    =   6975
   ClientLeft      =   150
   ClientTop       =   0
   ClientWidth     =   9840
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   465
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   656
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   2220
      TabIndex        =   13
      Text            =   "1"
      Top             =   4485
      Width           =   1305
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Put hole in Alpha"
      Height          =   240
      Left            =   225
      TabIndex        =   11
      Top             =   3825
      Width           =   1680
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   210
      TabIndex        =   5
      Text            =   "1"
      Top             =   4485
      Width           =   1560
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Process XP-Alpha.bmp"
      Height          =   525
      Left            =   195
      TabIndex        =   4
      Top             =   3135
      Width           =   1425
   End
   Begin VB.PictureBox picORG 
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   495
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   3
      Top             =   2295
      Width           =   720
   End
   Begin VB.PictureBox PICC 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   2250
      Left            =   4200
      Picture         =   "Main.frx":0000
      ScaleHeight     =   150
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   2
      Top             =   3855
      Width           =   2250
   End
   Begin VB.PictureBox PICA 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   6450
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   1
      Top             =   360
      Width           =   720
   End
   Begin VB.PictureBox PIC 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   2730
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   0
      Top             =   360
      Width           =   720
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Alpha divider"
      Height          =   225
      Index           =   1
      Left            =   2235
      TabIndex        =   12
      Top             =   4215
      Width           =   1275
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Loaded XP-Alpha"
      Height          =   240
      Index           =   3
      Left            =   105
      TabIndex        =   10
      Top             =   1950
      Width           =   1680
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Reconstruction"
      Height          =   240
      Index           =   2
      Left            =   4215
      TabIndex        =   9
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Alpha"
      Height          =   240
      Index           =   1
      Left            =   6495
      TabIndex        =   8
      Top             =   45
      Width           =   1365
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Image"
      Height          =   240
      Index           =   0
      Left            =   2760
      TabIndex        =   7
      Top             =   60
      Width           =   705
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Image multiplier"
      Height          =   225
      Index           =   0
      Left            =   330
      TabIndex        =   6
      Top             =   4215
      Width           =   1275
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Demonstration of a 32 bpp bmp from
' first principles.

' Alternatively a DIB method could be used to
' load the BMP using the LoadImage API.


Option Explicit

Private Type BITMAPINFOHEADER ' 40 bytes
   biSize As Long
   biwidth As Long
   biheight As Long
   biPlanes As Integer
   biBitCount As Integer
   biCompression As Long
   biSizeImage As Long
   biXPelsPerMeter As Long
   biYPelsPerMeter As Long
   biClrUsed As Long
   biClrImportant As Long
End Type
Private BInfo As BITMAPINFOHEADER

Private Declare Function StretchDIBits Lib "gdi32.dll" _
   (ByVal hdc As Long, _
   ByVal x As Long, ByVal y As Long, _
   ByVal DX As Long, ByVal DY As Long, _
   ByVal SrcX As Long, ByVal SrcY As Long, _
   ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, _
   ByRef lpBits As Any, _
   ByRef BInfo As BITMAPINFOHEADER, _
   ByVal wUsage As Long, _
   ByVal dwRop As Long) As Long
   
Private Declare Function SetStretchBltMode Lib "gdi32.dll" _
(ByVal hdc As Long, ByVal nStretchMode As Long) As Long

Private Declare Function GetStretchBltMode Lib "gdi32.dll" _
   (ByVal hdc As Long) As Long

Const HALFTONE As Long = 4
Const COLORONCOLOR As Long = 3

Private Declare Function SetPixelV Lib "gdi32.dll" _
(ByVal hdc As Long, ByVal x As Long, ByVal y As Long, _
   ByVal crColor As Long) As Long

Private BMPSpec$          ' Fixed in this demo
Private BM$               ' BM
Private picARR() As Byte  ' Image
Private alfARR() As Byte  ' Alpha

Private addARR() As Byte  ' Reconstructed

Private W As Long, H As Long  ' Image width & height
Private mul As Single         ' Image multiplier
Private div As Single         ' Alpha divider
Private bpp As Integer        ' Bits per pixel
Private numplanes As Integer  ' 1

Private Sub Command1_Click()
'Seek [#]filenumber, position ' goes to location position in file

'BMP Header for 32 bpp
' BM                       ' 2 bytes or Integer 66+256*77 = 19778
' FileSize                 ' Long
' Reserved                 ' Integer
' Reserved                 ' Integer
' Offset to image data     ' Long (54)
' BitmapInfo Header size   ' Long (40)
' Width                    ' Long   ' Location 19
' Height                   ' Long
' Planes                   ' Integer  (1)
' BitsPerPixel             ' Integer  (32)
' 6 Longs                  ' Compression etc
' Location 55  image data  BGRA,BGRA etc
'                       ie Blue,Green,Red,Alpha bytes (0-255)

' Picboxes:-
' PICORG  loaded image
' PIC     image
' PICA    alpha bytes
' PICC    reconstructed transparent image

Dim fnum As Long
Dim BMPsize As Long
Dim ix As Long, iy As Long
Dim Alf As Byte    ' Alpha byte 0 - 255
Dim fac As Single  ' Alpha multiplier
Dim R As Byte, G As Byte, B As Byte
Dim RB As Byte, GB As Byte, BB As Byte
Dim Cul As Long
   'mul = 4  ' Image multiplier
   On Error GoTo BMPError
   
   Cls
   '==================================
   ' Read info directly from file
   fnum = FreeFile
   Open BMPSpec$ For Binary As fnum
   BM$ = Input(2, #fnum)
   
   If BM$ <> "BM" Then
      MsgBox "Not a BMP file"
      Close fnum
      Exit Sub
   End If
   
   ' Alternatively use a Type structure for this
   
   Seek #fnum, 19    ' Location of Width
   Get #fnum, , W
   Get #fnum, , H
   Get #fnum, , numplanes
   Get #fnum, , bpp
   
   If bpp <> 32 Then
      MsgBox "Not 32 bpp"
      Close fnum
      Exit Sub
   Else
      ReDim picARR(0 To 3, W - 1, H - 1)
      Seek #fnum, 55    ' Location of BGRA bytes
      Get #fnum, , picARR()
   End If
   Close fnum
   '==================================
   ' Image size
   BMPsize = W * 4 * H
   ' Size picboxes
   With PIC    ' Image
      .Cls
      .Width = mul * W
      .Height = mul * H
      .Refresh
   End With
   With PICA   ' Alpha
      .Cls
      .Width = mul * W
      .Height = mul * H
      .Refresh
   End With
'   Could resize PICC :-
'   With PICC   ' Reconstructed
'      .Cls
'      .Width = mul * W
'      .Height = mul * H
'      .Refresh
'   End With
   '==================================
   ' Show some info
   Print " " & BM$
   Print " No. pic bytes= "; BMPsize
   Print " W= "; W
   Print " H= "; H
   Print " bpp= "; bpp
   Print " Image mul= "; mul
   Print " Alpha  div= "; div
   '==================================
   ' Transfer alpha bytes to alfARR()
   ReDim alfARR(0 To 3, W - 1, H - 1)
   For iy = 0 To H - 1
   For ix = 0 To W - 1
      Alf = picARR(3, ix, iy) / div  '/ >1 image more transparent
      alfARR(0, ix, iy) = Alf
      alfARR(1, ix, iy) = Alf
      alfARR(2, ix, iy) = Alf
   Next ix
   Next iy
   
   '==================================
   ' Show image and alpha-mask
   With BInfo
      .biSize = 40
      .biwidth = W
      .biheight = H
      .biPlanes = 1
      .biBitCount = 32
   End With
   
   SetStretchBltMode PIC.hdc, HALFTONE
   
   StretchDIBits PIC.hdc, 0, 0, mul * W, mul * H, _
      0, 0, W, H, picARR(0, 0, 0), _
      BInfo, 0, vbSrcCopy
   
   PIC.Refresh
   
   SetStretchBltMode PICA.hdc, COLORONCOLOR
   
   StretchDIBits PICA.hdc, 0, 0, mul * W, mul * H, _
      0, 0, W, H, alfARR(0, 0, 0), _
      BInfo, 0, vbSrcCopy
   
   PICA.Refresh
   
   '==================================
   ' Put transparent hole in alpha
   If Check1.Value = Checked Then
      Label2(1) = "Alpha + hole"
      PICA.FillStyle = 0
      PICA.Circle (mul * W / 2, mul * H / 2), mul * 8
      PICA.FillStyle = 1
   Else
      Label2(1) = "Alpha"
   End If

   '==================================
   ' Show image with alpha
   ReDim addARR(0 To 3, mul * W - 1, mul * H - 1)
   PICC.Cls
   For iy = 0 To mul * H - 1
   For ix = 0 To mul * W - 1
      ' Backgound
      Cul = PICC.Point(ix, iy)   ' NB Point faster than GetPixel !
      RB = Cul And &HFF&
      GB = (Cul And &HFF00&) \ &H100&
      BB = (Cul And &HFF0000) \ &H10000
      ' Image
      Cul = PIC.Point(ix, iy)
      R = Cul And &HFF&
      G = (Cul And &HFF00&) \ &H100&
      B = (Cul And &HFF0000) \ &H10000
      ' Alpha
      Cul = PICA.Point(ix, iy)
      Alf = Cul And &HFF&
      fac = Alf / 255
      
      ' fac = 0  B=BB  background
      ' fac = 1  B=B   image
      B = BB * (1 - fac) + B * fac
      G = GB * (1 - fac) + G * fac
      R = RB * (1 - fac) + R * fac
   
      addARR(0, ix, iy) = B
      addARR(1, ix, iy) = G
      addARR(2, ix, iy) = R
      
      ' Or
      'PICC.PSet (ix, iy ), RGB(R, G, B)
      'SetPixelV PICC.hdc, ix, iy, RGB(R, G, B)
   Next ix
   Next iy
   
   With BInfo
      .biSize = 40
      .biwidth = mul * W
      .biheight = -mul * H
      .biPlanes = 1
      .biBitCount = 32
   End With
   'SetStretchBltMode PICC.hdc, HALFTONE
   StretchDIBits PICC.hdc, 0, 0, mul * W, mul * H, _
      0, 0, mul * W, mul * H, addARR(0, 0, 0), _
      BInfo, 0, vbSrcCopy
   
   Exit Sub
   '==========
BMPError:
   MsgBox "FILE ERROR"
   Close
End Sub

Private Sub Combo1_Click()
' Image multiplier
   mul = 1 + Val(Combo1.ListIndex)
End Sub

Private Sub Combo2_Click()
' Alpha divider
   div = (2 + Val(Combo2.ListIndex)) * 0.5
End Sub

Private Sub Form_Load()
   ' One off test bmp
   BMPSpec$ = "XP-Alpha.bmp"
   With picORG
      .Width = 48
      .Height = 48
   End With
   picORG.Picture = LoadPicture(BMPSpec$)
   picORG.Refresh
   mul = 1
   Combo1.AddItem "1"
   Combo1.AddItem "2"
   Combo1.AddItem "3"
   Combo1.AddItem "4"
   Combo1.ListIndex = 0
   
   Combo2.AddItem "1"
   Combo2.AddItem "1.5"
   Combo2.AddItem "2"
   Combo2.AddItem "2.5"
   Combo2.ListIndex = 0

   With PICC   ' For Reconstructed image
      .Cls
      .Width = 192
      .Height = 192
      .Refresh
   End With

End Sub
