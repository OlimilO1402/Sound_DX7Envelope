VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "DX7 Envelope Generator EG"
   ClientHeight    =   10575
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13815
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   705
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   921
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command6 
      Caption         =   "Def"
      Height          =   375
      Left            =   2640
      TabIndex        =   46
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Play"
      Height          =   375
      Left            =   5160
      TabIndex        =   45
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Load Env.html"
      Height          =   375
      Left            =   3480
      TabIndex        =   41
      Top             =   120
      Width           =   1575
   End
   Begin VB.HScrollBar HScroll0 
      Height          =   255
      Left            =   6360
      Max             =   0
      Min             =   99
      TabIndex        =   38
      Top             =   1080
      Value           =   99
      Width           =   1815
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   255
      Left            =   6360
      Max             =   0
      Min             =   99
      TabIndex        =   37
      Top             =   2160
      Value           =   99
      Width           =   1815
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      Left            =   6360
      Max             =   0
      Min             =   99
      TabIndex        =   35
      Top             =   1800
      Value           =   99
      Width           =   1815
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   6360
      Max             =   0
      Min             =   99
      TabIndex        =   33
      Top             =   1440
      Value           =   99
      Width           =   1815
   End
   Begin VB.VScrollBar VScroll3 
      Height          =   1815
      Left            =   5520
      Max             =   0
      Min             =   99
      TabIndex        =   31
      Top             =   840
      Value           =   99
      Width           =   255
   End
   Begin VB.VScrollBar VScroll2 
      Height          =   1815
      Left            =   5160
      Max             =   0
      Min             =   99
      TabIndex        =   29
      Top             =   840
      Value           =   99
      Width           =   255
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1815
      Left            =   4800
      Max             =   0
      Min             =   99
      TabIndex        =   28
      Top             =   840
      Value           =   99
      Width           =   255
   End
   Begin VB.VScrollBar VScroll0 
      Height          =   1815
      Left            =   4440
      Max             =   0
      Min             =   99
      TabIndex        =   25
      Top             =   840
      Value           =   99
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Bsp3"
      Height          =   375
      Left            =   1800
      TabIndex        =   24
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Bsp2"
      Height          =   375
      Left            =   960
      TabIndex        =   23
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Bsp1"
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox param7 
      Alignment       =   1  'Rechts
      Height          =   315
      Left            =   3000
      TabIndex        =   18
      Text            =   "0"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox param5 
      Alignment       =   1  'Rechts
      Height          =   315
      Left            =   3000
      TabIndex        =   16
      Text            =   "80"
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox param3 
      Alignment       =   1  'Rechts
      Height          =   315
      Left            =   840
      TabIndex        =   14
      Text            =   "0"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox param1 
      Alignment       =   1  'Rechts
      Height          =   315
      Left            =   840
      TabIndex        =   12
      Text            =   "0"
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox nsamp 
      Alignment       =   1  'Rechts
      Height          =   315
      Left            =   1920
      TabIndex        =   10
      Text            =   "4000"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox param6 
      Alignment       =   1  'Rechts
      Height          =   315
      Left            =   3000
      TabIndex        =   8
      Text            =   "0"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox param4 
      Alignment       =   1  'Rechts
      Height          =   315
      Left            =   3000
      TabIndex        =   6
      Text            =   "80"
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox param2 
      Alignment       =   1  'Rechts
      Height          =   315
      Left            =   840
      TabIndex        =   4
      Text            =   "0"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox param0 
      Alignment       =   1  'Rechts
      Height          =   315
      Left            =   840
      TabIndex        =   2
      Text            =   "99"
      Top             =   600
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      FillColor       =   &H00E0E0E0&
      Height          =   7575
      Left            =   120
      ScaleHeight     =   505
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   905
      TabIndex        =   0
      Top             =   2760
      Width           =   13575
   End
   Begin VB.Label Label21 
      Caption         =   "Label21"
      Height          =   255
      Left            =   9480
      TabIndex        =   44
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label22 
      Caption         =   "Sampleformat of a DX7 is 12bit 49.096kHz"
      Height          =   495
      Left            =   8280
      TabIndex        =   43
      Top             =   1200
      Width           =   3855
   End
   Begin VB.Label LblDuration 
      Caption         =   "Duration:"
      Height          =   255
      Left            =   1920
      TabIndex        =   42
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Label Label20 
      Caption         =   "R4"
      Height          =   255
      Left            =   6000
      TabIndex        =   40
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label Label19 
      Caption         =   "R3"
      Height          =   255
      Left            =   6000
      TabIndex        =   39
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label18 
      Caption         =   "R2"
      Height          =   255
      Left            =   6000
      TabIndex        =   36
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Label17 
      Caption         =   "R1"
      Height          =   255
      Left            =   6000
      TabIndex        =   34
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label16 
      Caption         =   "L4"
      Height          =   255
      Left            =   5520
      TabIndex        =   32
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label15 
      Caption         =   "L3"
      Height          =   255
      Left            =   5160
      TabIndex        =   30
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label14 
      Caption         =   "L2"
      Height          =   255
      Left            =   4800
      TabIndex        =   27
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label13 
      Caption         =   "L1"
      Height          =   255
      Left            =   4440
      TabIndex        =   26
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label12 
      Caption         =   "Similarly for rate, this time measured in dB/s, for decay. Attack is faster, and has a nonlinear curve."
      Height          =   495
      Left            =   8280
      TabIndex        =   21
      Top             =   600
      Width           =   3855
   End
   Begin VB.Label Label11 
      Caption         =   $"Form1.frx":1782
      Height          =   495
      Left            =   8280
      TabIndex        =   20
      Top             =   120
      Width           =   5415
   End
   Begin VB.Label Label10 
      Caption         =   "Envelope:"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label9 
      Caption         =   "Rate 4:"
      Height          =   255
      Left            =   2280
      TabIndex        =   17
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "Rate 3:"
      Height          =   255
      Left            =   2280
      TabIndex        =   15
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "Rate 2:"
      Height          =   255
      Left            =   2280
      TabIndex        =   13
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Rate 1:"
      Height          =   255
      Left            =   2280
      TabIndex        =   11
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Number of samples:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Level 4:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Level 3:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Level 2:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Level 1:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Env As Env
Dim bInit    As Boolean
'the sampleformat of a DX7 is 12bit 49.096kHz

Private Sub Form_Load()
    Command2_Click
End Sub

Private Sub Command1_Click()
    bInit = True
    param0.Text = 99:     param4.Text = 80
    param1.Text = 0:      param5.Text = 80
    param2.Text = 0:      param6.Text = 0
    param3.Text = 0:      param7.Text = 0
    nsamp.Text = 4000
    UpdateScrolls
    bInit = False
    redraw
End Sub

Private Sub Command2_Click()
    bInit = True
    param0.Text = 99:     param4.Text = 80
    param1.Text = 60:     param5.Text = 82
    param2.Text = 33:     param6.Text = 75
    param3.Text = 0:      param7.Text = 80
    nsamp.Text = 6000
    UpdateScrolls
    bInit = False
    redraw
End Sub

Private Sub Command3_Click()
    bInit = True
    param0.Text = 99:     param4.Text = 80
    param1.Text = 22:     param5.Text = 78
    param2.Text = 77:     param6.Text = 51
    param3.Text = 0:      param7.Text = 83
    nsamp.Text = 10000
    UpdateScrolls
    bInit = False
    redraw
End Sub

Private Sub Command5_Click()
    'Dim Env As Env: Set Env = New_Env(GetParams())
    MsgBox "niy=not implemented yet"
End Sub

Private Sub Command6_Click()
    bInit = True
    param0.Text = 99:     param4.Text = 99
    param1.Text = 99:     param5.Text = 99
    param2.Text = 99:     param6.Text = 99
    param3.Text = 0:      param7.Text = 99
    nsamp.Text = 10000
    UpdateScrolls
    bInit = False
    redraw
End Sub


Sub UpdateScrolls()
    VScroll0.Value = param0.Text:      HScroll0.Value = param4.Text
    VScroll1.Value = param1.Text:      HScroll1.Value = param5.Text
    VScroll2.Value = param2.Text:      HScroll2.Value = param6.Text
    VScroll3.Value = param3.Text:      HScroll3.Value = param7.Text
End Sub

Private Sub Command4_Click()
    'Shell "explorer.exe " & App.Path
    'Shell "explorer.exe " & App.Path & "\Resources\env_jscript.html"
    LoadEnvHtml GetParams(), GetNSamples
End Sub

Private Sub Form_Resize()
    Dim l As Single: l = Picture1.Left
    Dim t As Single: t = Picture1.Top
    Dim brdr As Single: brdr = l
    Dim w As Single: w = Me.ScaleWidth - 2 * brdr
    Dim h As Single: h = Me.ScaleHeight - t - brdr
    If w > 0 And h > 0 Then Picture1.Move l, t, w, h
    redraw
End Sub

Private Sub param0_LostFocus(): VScroll0.Value = param0.Text: redraw: End Sub
Private Sub param1_LostFocus(): VScroll1.Value = param1.Text: redraw: End Sub
Private Sub param2_LostFocus(): VScroll2.Value = param2.Text: redraw: End Sub
Private Sub param3_LostFocus(): VScroll3.Value = param3.Text: redraw: End Sub
Private Sub param4_LostFocus(): HScroll0.Value = param4.Text: redraw: End Sub
Private Sub param5_LostFocus(): HScroll1.Value = param5.Text: redraw: End Sub
Private Sub param6_LostFocus(): HScroll2.Value = param6.Text: redraw: End Sub
Private Sub param7_LostFocus(): HScroll3.Value = param7.Text: redraw: End Sub
Private Sub nsamp_LostFocus():    redraw: End Sub

Private Sub VScroll0_Change(): param0.Text = VScroll0.Value: redraw: End Sub
Private Sub VScroll1_Change(): param1.Text = VScroll1.Value: redraw: End Sub
Private Sub VScroll2_Change(): param2.Text = VScroll2.Value: redraw: End Sub
Private Sub VScroll3_Change(): param3.Text = VScroll3.Value: redraw: End Sub

Private Sub HScroll0_Change(): param4.Text = HScroll0.Value: redraw: End Sub
Private Sub HScroll1_Change(): param5.Text = HScroll1.Value: redraw: End Sub
Private Sub HScroll2_Change(): param6.Text = HScroll2.Value: redraw: End Sub
Private Sub HScroll3_Change(): param7.Text = HScroll3.Value: redraw: End Sub

Sub redraw()
    If bInit Then Exit Sub
    Dim params() As Long: params = GetParams
    Dim nsamples As Long: nsamples = GetNSamples
    LblDuration.Caption = "Duration: " & Format(nsamples / 49096, "0.000") & " sec"
    Dim data() As Long: data = envdata(params, nsamples)
    chart_draw data, "Envelope" 'chart.draw(data, {title: 'Envelope'})
End Sub

Function GetParams() As Long()
    ReDim params(0 To 7) As Long
    Dim i As Long
    For i = 0 To 7
        params(i) = CLng(Form1.Controls("param" & i).Text)
    Next
    GetParams = params
End Function

Function GetNSamples() As Long
    GetNSamples = CLng(Form1.Controls("nsamp").Text)
End Function
    
Sub chart_draw(data() As Long, title As String)
    'draw the underlying grid
    Picture1.Cls
    Dim fcol As Long: fcol = Picture1.ForeColor
    'Picture1.ForeColor = vbBlack
    Dim BrdrX As Long: BrdrX = 20
    Dim BrdrY As Long: BrdrY = BrdrX
    Dim u As Long: u = UBound(data)
    Dim sh As Long: sh = Picture1.ScaleHeight - 2 * BrdrY
    Dim sw As Long: sw = Picture1.ScaleWidth - 2 * BrdrX
    Dim ZeroX As Long: ZeroX = BrdrX
    Dim ZeroY As Long: ZeroY = BrdrY + sh
    Dim dy As Double: dy = sh / 4000  '/ maximum level is 4000 (=40dB)
    Dim dx As Double: dx = sw / (u + 1)
    Dim X1 As Long, Y1 As Long
    Dim X2 As Long, Y2 As Long
    Dim i As Long
    'prepare the Envelope
    Dim maxy As Long
    ReDim polyl(0 To u) As pointapi
    For i = 0 To u
        polyl(i).X = ZeroX + i * dx
        polyl(i).Y = ZeroY - data(i) * dy
        maxy = Max(maxy, data(i))
    Next
    'draw hor lines
    X1 = ZeroX
    X2 = ZeroX + u * dx
    For i = 0 To 4
        Y1 = ZeroY - i * 1000 * dy
        Y2 = Y1
        Picture1.Line (X1, Y1)-(X2, Y2)
    Next
    'draw vert lines
    Y1 = ZeroY
    Y2 = BrdrY
    For i = 0 To u / 1000
        X1 = ZeroX + i * 1000 * dx
        X2 = X1
        Picture1.Line (X1, Y1)-(X2, Y2)
    Next
    'draw the Envelope
    Picture1.ForeColor = vbBlue
    Polyline Picture1.hdc, polyl(0), u + 1
    'draw NoteOff
    X1 = ZeroX + Env.NoteOff_X * dx
    Y1 = ZeroY - Env.NoteOff_Y * dy
    Picture1.Circle (X1, Y1), 3
    Label21.Caption = maxy
    Picture1.ForeColor = fcol
End Sub

'function envdata(params, nsamp) {
'  console.log(nsamp);
'  var result = [['samp', 'env']];
'  var env = new Env(params);
'  for (var i = 0; i < nsamp; i++) {
'    if (i == 3 * nsamp / 4) {
'      env.keyup();
'    }
'    result.push([i, env.getsample()]);
'  }
'  return result;
'}
Public Function envdata(params() As Long, ByVal nsamp As Long) As Long() 'Variant()
    'Debug.Print nsamp 'console.Log (nsamp)
    Dim i As Long
    ReDim result(0 To nsamp - 1) As Long
    'Dim Env As Env:
    Set Env = New_Env(params)
    For i = 0 To nsamp - 1
        If (i = 3 * nsamp / 4) Then
            Env.keyup
        End If
        result(i) = Env.getsample
    Next
    envdata = result
End Function

