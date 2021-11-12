VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Demo - PI engine, digits to signals"
   ClientHeight    =   10650
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19845
   LinkTopic       =   "Form1"
   ScaleHeight     =   710
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1323
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Orig 
      Height          =   1215
      Index           =   0
      Left            =   2640
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   35
      Text            =   "pi.frx":0000
      Top             =   11280
      Width           =   10335
   End
   Begin VB.TextBox Orig 
      Height          =   1215
      Index           =   1
      Left            =   2640
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   34
      Text            =   "pi.frx":2727
      Top             =   12720
      Width           =   10335
   End
   Begin VB.TextBox Orig 
      Height          =   1215
      Index           =   2
      Left            =   2640
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   33
      Text            =   "pi.frx":4E4E
      Top             =   14160
      Width           =   10335
   End
   Begin VB.TextBox Orig 
      Height          =   1215
      Index           =   3
      Left            =   2640
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   32
      Text            =   "pi.frx":7575
      Top             =   15600
      Width           =   10335
   End
   Begin VB.Frame Frame3 
      Caption         =   "Analyze the digits from:"
      Height          =   2655
      Left            =   240
      TabIndex        =   26
      Top             =   120
      Width           =   8775
      Begin VB.TextBox Sec1 
         Height          =   2055
         Left            =   1200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   40
         Top             =   360
         Width           =   7335
      End
      Begin VB.OptionButton Option2 
         Caption         =   "PI"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   30
         Top             =   480
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton Option2 
         Caption         =   "e"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   29
         Top             =   720
         Width           =   495
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Ratio"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   28
         Top             =   960
         Width           =   735
      End
      Begin VB.OptionButton Option2 
         Caption         =   "sqr(2)"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   27
         Top             =   1200
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Make signal by rule:"
      Height          =   1815
      Left            =   14640
      TabIndex        =   20
      Top             =   7920
      Width           =   5055
      Begin VB.CheckBox OSignal 
         Caption         =   "Overlapping signals"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   1320
         Width           =   1935
      End
      Begin VB.HScrollBar x_tuple 
         Height          =   255
         Left            =   1800
         Max             =   20
         Min             =   1
         TabIndex        =   24
         Top             =   600
         Value           =   1
         Width           =   3015
      End
      Begin VB.OptionButton Option1 
         Caption         =   "[n-1<n]"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   23
         ToolTipText     =   "If [n-1<n] then add one to the signal else subtract one"
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "[n-1>n]"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   22
         ToolTipText     =   "If [n-1>n] then add one to the signal else subtract one"
         Top             =   480
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Compare digits"
         Height          =   495
         Left            =   2400
         TabIndex        =   21
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label Arata_tuple 
         BackStyle       =   0  'Transparent
         Caption         =   "n-tuple"
         Height          =   255
         Left            =   1800
         TabIndex        =   25
         Top             =   360
         Width           =   3015
      End
   End
   Begin VB.PictureBox MaxFWindow 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   9360
      ScaleHeight     =   103
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   335
      TabIndex        =   19
      Top             =   7920
      Width           =   5055
   End
   Begin VB.PictureBox ProbGraph 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   14640
      ScaleHeight     =   103
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   335
      TabIndex        =   18
      Top             =   360
      Width           =   5055
   End
   Begin VB.OptionButton LineSum2 
      Caption         =   "Cross Line"
      Height          =   255
      Left            =   17040
      TabIndex        =   17
      Top             =   4200
      Width           =   1095
   End
   Begin VB.OptionButton CubeSum 
      Caption         =   "Cube Line"
      Height          =   255
      Left            =   15960
      TabIndex        =   15
      Top             =   4200
      Width           =   1575
   End
   Begin VB.OptionButton BarSum 
      Caption         =   "Cross Bar"
      Height          =   255
      Left            =   18120
      TabIndex        =   14
      Top             =   4200
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.OptionButton LineSum 
      Caption         =   "Line"
      Height          =   195
      Left            =   15240
      TabIndex        =   13
      Top             =   4200
      Width           =   975
   End
   Begin VB.PictureBox SumOfPI 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   14640
      ScaleHeight     =   103
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   335
      TabIndex        =   9
      Top             =   2280
      Width           =   5055
   End
   Begin VB.PictureBox DigitFrec 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   9360
      ScaleHeight     =   111
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   335
      TabIndex        =   8
      Top             =   5640
      Width           =   5055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Scan digits:"
      Height          =   3135
      Left            =   14640
      TabIndex        =   3
      Top             =   4680
      Width           =   5055
      Begin VB.CommandButton Stop_Run 
         Caption         =   "Stop"
         Height          =   495
         Left            =   2640
         TabIndex        =   16
         Top             =   2400
         Width           =   1935
      End
      Begin VB.HScrollBar Stepp 
         Height          =   255
         Left            =   240
         Max             =   100
         Min             =   1
         TabIndex        =   7
         Top             =   1200
         Value           =   10
         Width           =   4695
      End
      Begin VB.HScrollBar pause 
         Height          =   255
         Left            =   240
         Max             =   100
         Min             =   5
         TabIndex        =   6
         Top             =   1920
         Value           =   10
         Width           =   4695
      End
      Begin VB.HScrollBar Window 
         Height          =   255
         Left            =   240
         Max             =   5000
         Min             =   20
         TabIndex        =   5
         Top             =   600
         Value           =   100
         Width           =   4695
      End
      Begin VB.CommandButton Anim 
         Caption         =   "Run window"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   4
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label StepTxT 
         BackStyle       =   0  'Transparent
         Caption         =   "Sliding window step:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Processing speed:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Label WindowTXT 
         BackStyle       =   0  'Transparent
         Caption         =   "Sliding window length:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.PictureBox Center_patt 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   9360
      ScaleHeight     =   311
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   335
      TabIndex        =   2
      Top             =   600
      Width           =   5055
   End
   Begin VB.CommandButton TR_PI 
      Caption         =   "Process all in one"
      Height          =   615
      Left            =   1920
      TabIndex        =   1
      Top             =   9120
      Width           =   4935
   End
   Begin VB.TextBox WMatrix 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   3240
      Width           =   8775
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   -5400
      Picture         =   "pi.frx":9C9C
      Top             =   9960
      Width           =   25290
   End
   Begin VB.Line Line2 
      X1              =   976
      X2              =   976
      Y1              =   260
      Y2              =   240
   End
   Begin VB.Line Line1 
      X1              =   1312
      X2              =   1312
      Y1              =   245
      Y2              =   260
   End
   Begin VB.Label E_SW 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      Height          =   255
      Left            =   17880
      TabIndex        =   48
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label B_SW 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Left            =   14580
      TabIndex        =   47
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Results in text format:"
      Height          =   255
      Left            =   240
      TabIndex        =   46
      Top             =   3000
      Width           =   4935
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Balance on sliding windows:"
      Height          =   255
      Left            =   14640
      TabIndex        =   45
      Top             =   2040
      Width           =   4935
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Parallel signals for each digit along the sequence"
      Height          =   255
      Left            =   14640
      TabIndex        =   44
      Top             =   120
      Width           =   4935
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Max digit frequency on all sliding windows:"
      Height          =   255
      Left            =   9360
      TabIndex        =   43
      Top             =   7680
      Width           =   4935
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Digit frequency on sliding window/all digits:"
      Height          =   255
      Left            =   9360
      TabIndex        =   42
      Top             =   5400
      Width           =   4935
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Transition matrix between the digits of the sliding window/all digits:"
      Height          =   255
      Left            =   9240
      TabIndex        =   41
      Top             =   120
      Width           =   4935
   End
   Begin VB.Label Label6 
      Caption         =   "Pi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   960
      TabIndex        =   39
      Top             =   11640
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "e"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   960
      TabIndex        =   38
      Top             =   12960
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "sqr(2)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   960
      TabIndex        =   37
      Top             =   15840
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Ratio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   960
      TabIndex        =   36
      Top             =   14400
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   ________________________________                          ____________________
'  /          PI Engine             \________________________/       v1           |
' |                                                                               |
' |            Name:  PI Engine                                                   |
' |        Category:  Open source software                                        |
' |          Author:  Paul A. Gagniuc                                             |
' |            Book:  Algorithms in Bioinformatics: Theory and Implementation     |
' |           Email:  paul_gagniuc@acad.ro                                        |
' |  ____________________________________________________________________________ |
' |                                                                               |
' |    Date Created:  May 2014                                                    |
' |          Update:  August 2021                                                 |
' |       Tested On:  WinXP, WinVista, Win7, Win8, Win10                          |
' |             Use:  Analysis                                                    |
' |                                                                               |
' |                  _____________________________                                |
' |_________________/                             \_______________________________|

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim stop_anim As Boolean
'=IF(E2>F2,D8+1,D8-1)

Dim M1(0 To 9, 0 To 9) As String
Dim MatrixMaxVal As Variant 'matrix maximum value
Dim MatrixMinVal As Variant 'matrix minimum value

Dim MAX_DIGIT_WINDOW(0 To 9) As Integer

Private Sub Anim_Click()

    stop_anim = False
    
    Sec1.Text = Replace(Sec1.Text, vbCrLf, "")
    
    For i = 0 To 9
        MAX_DIGIT_WINDOW(i) = 0
    Next i
    
    For s = 1 To Len(Sec1.Text) - Window.Value Step Stepp.Value
        
        Call Fill_Transition_Matrix(M1, 9, 9, Mid(Sec1.Text, s, Window.Value))
        Call frec(Mid(Sec1.Text, s, Window.Value))
        Call SumPI2(Mid(Sec1.Text, s, Window.Value))
        Call Max_frec(Mid(Sec1.Text, s, Window.Value))
        
        Call P(Mid(Sec1.Text, s, Window.Value), s)
        
        Sleep (CLng(pause.Value))
        
        '---------------------------------------
        MatrixMaxVal = 0
        MatrixMinVal = 0
        
        For i = 0 To 9
            For j = 0 To 9
                    If M1(i, j) > MatrixMaxVal Then MatrixMaxVal = M1(i, j)
                    If M1(i, j) < MatrixMinVal Then MatrixMinVal = M1(i, j)
            Next j
        Next i
        
        Call DrowColorMatrix(9, 9, M1, Center_patt)
        '---------------------------------------
        Frame1.Caption = "Scan digits: " & (s + Window.Value + Stepp.Value - 1)
        
        E_SW.Caption = (s + Window.Value + Stepp.Value - 1)
        B_SW.Caption = E_SW.Caption - Window.Value
        
        
        If stop_anim = True Then GoTo 1
        DoEvents
    
    Next s
1:
End Sub


Private Sub Command1_Click()

    If Option1(0).Value = True Then a = 0
    If Option1(1).Value = True Then a = 1
    
    WMatrix.Text = WMatrix.Text & vbCrLf & RulePI(Sec1.Text, a)
    WMatrix.SelStart = Len(WMatrix.Text)
    
End Sub



Function P(ByVal s As String, ByVal pos As Integer)
    'd = cati de x sunt in fereastra
    
    Dim f(0 To 9) As Integer
    Dim prob(0 To 9) As Variant
    
    For i = 1 To Len(s)
    
        DI = Mid(s, i, 1)
    
        If DI = "0" Then f(0) = f(0) + 1
        If DI = "1" Then f(1) = f(1) + 1
        If DI = "2" Then f(2) = f(2) + 1
        If DI = "3" Then f(3) = f(3) + 1
        If DI = "4" Then f(4) = f(4) + 1
        If DI = "5" Then f(5) = f(5) + 1
        If DI = "6" Then f(6) = f(6) + 1
        If DI = "7" Then f(7) = f(7) + 1
        If DI = "8" Then f(8) = f(8) + 1
        If DI = "9" Then f(9) = f(9) + 1
    
    Next i
    
    
    
    peX = (ProbGraph.ScaleWidth / Len(Sec1.Text))
    peY = (ProbGraph.ScaleHeight / 100)
    
    old = 0
    
    For i = 0 To 9
    
        prob(i) = f(i) / Len(s)
        nou_p = prob(i) * 100
        
        tmp = old + nou_p
        
        If i = 0 Then colora = RGB(55, 55, 55)
        If i = 1 Then colora = RGB(0, 55, 55)
        If i = 2 Then colora = RGB(55, 0, 55)
        If i = 3 Then colora = RGB(55, 70, 55)
        If i = 4 Then colora = RGB(0, 55, 100)
        If i = 5 Then colora = RGB(55, 160, 55)
        If i = 6 Then colora = RGB(55, 55, 200)
        If i = 7 Then colora = RGB(0, 0, 55)
        If i = 8 Then colora = RGB(55, 0, 0)
        If i = 9 Then colora = RGB(55, 90, 160)
        
        
        If i = 0 Then colora = RGB(i * 30, 55, 55)
        If i = 1 Then colora = RGB(i * 30, 55, 55)
        If i = 2 Then colora = RGB(i * 30, 55, 55)
        If i = 3 Then colora = RGB(i * 30, 55, 55)
        If i = 4 Then colora = RGB(i * 30, 55, 55)
        If i = 5 Then colora = RGB(i * 30, 55, 55)
        If i = 6 Then colora = RGB(i * 30, 55, 55)
        If i = 7 Then colora = RGB(i * 30, 55, 55)
        If i = 8 Then colora = RGB(i * 30, 55, 55)
        If i = 9 Then colora = RGB(i * 30, 55, 55)
        
        If i = 0 Then colora = RGB(i * 30, i * 20, 55)
        If i = 1 Then colora = RGB(i * 30, i * 20, 55)
        If i = 2 Then colora = RGB(i * 30, i * 20, 55)
        If i = 3 Then colora = RGB(i * 30, i * 20, 55)
        If i = 4 Then colora = RGB(i * 30, i * 20, 55)
        If i = 5 Then colora = RGB(i * 30, i * 20, 55)
        If i = 6 Then colora = RGB(i * 30, i * 20, 55)
        If i = 7 Then colora = RGB(i * 30, i * 20, 55)
        If i = 8 Then colora = RGB(i * 30, i * 20, 55)
        If i = 9 Then colora = RGB(i * 30, i * 20, 55)
        
        ProbGraph.Line (peX * pos, ProbGraph.ScaleHeight - (peY + old))-(peX * (pos + Len(s)), ProbGraph.ScaleHeight - (peY * tmp)), colora, BF
        'ProbGraph.Line (peX * pos, ProbGraph.ScaleHeight - (peY + old))-(peX * (pos), ProbGraph.ScaleHeight - (peY * tmp)), colora
        
        old = tmp
    
    Next i
    

End Function



Function SumPI2(ByVal s As String)

    a = 0
    
    For i = 1 To Len(s)
        DI = Val(Mid(s, i, 1))
        a = a + DI
    Next i
    
    peX = (SumOfPI.ScaleWidth / Len(s))
    peY = (SumOfPI.ScaleHeight / a)
    
    SumOfPI.Cls
    
    For i = 1 To Len(s)
    
        DI = Val(Mid(s, i, 1))
        d = d + DI
    
        If LineSum.Value = True Then SumOfPI.Line (peX * i, SumOfPI.ScaleHeight - (peY * old))-(peX * (i + 1), SumOfPI.ScaleHeight - (peY * d)), RGB(55, 55, 55)
        If LineSum2.Value = True Then SumOfPI.Line (peX * i, SumOfPI.ScaleHeight - (peY * old))-(peX * (i + 1), (peY * d)), RGB(55, 55, 55)
        If BarSum.Value = True Then SumOfPI.Line (peX * i, peY * old)-(peX * (i + 1), SumOfPI.ScaleHeight - (peY * d)), RGB(55, 55, 55), B
        'If CubeSum.Value = True Then SumOfPI.Line (peX * i, peY * old)-(peX * (i + 1), (peY * d)), RGB(55, 55, 55), B
        If CubeSum.Value = True Then SumOfPI.Line (peX * i, SumOfPI.ScaleHeight - (peY * old))-(peX * (i + 1), SumOfPI.ScaleHeight - (peY * d)), RGB(55, 55, 55), B
    
        old = d
    
    Next i

End Function




Function Max_frec(ByVal s As String)

    '------------------------------------
    Dim f(0 To 9) As Integer
    
    For i = 1 To Len(s)
    
        DI = Mid(s, i, 1)
    
        If DI = "0" Then f(0) = f(0) + 1
        If DI = "1" Then f(1) = f(1) + 1
        If DI = "2" Then f(2) = f(2) + 1
        If DI = "3" Then f(3) = f(3) + 1
        If DI = "4" Then f(4) = f(4) + 1
        If DI = "5" Then f(5) = f(5) + 1
        If DI = "6" Then f(6) = f(6) + 1
        If DI = "7" Then f(7) = f(7) + 1
        If DI = "8" Then f(8) = f(8) + 1
        If DI = "9" Then f(9) = f(9) + 1
    
    Next i
    
    For i = 0 To UBound(f)
        If f(i) > Maxim Then
            Maxim = f(i)
            tmp = i
        End If
    Next i
    
    MAX_DIGIT_WINDOW(tmp) = MAX_DIGIT_WINDOW(tmp) + 1
    
    peX = (MaxFWindow.ScaleWidth / (UBound(f) + 1))
    
    For i = 0 To UBound(MAX_DIGIT_WINDOW)
        If MAX_DIGIT_WINDOW(i) > Maxim Then Maxim = MAX_DIGIT_WINDOW(i)
    Next i
    
    peY = (MaxFWindow.ScaleHeight / Maxim)
    'peY = (MaxFWindow.ScaleHeight / Len(s))
    
    MaxFWindow.Cls
    
    For i = 0 To UBound(f)
        MaxFWindow.Line (peX * i, DigitFrec.ScaleHeight)-(peX * (i + 1), DigitFrec.ScaleHeight - (peY * MAX_DIGIT_WINDOW(i))), RGB(55, 20, 20), BF
        'MaxFWindow.Line (peX * i, DigitFrec.ScaleHeight)-(peX * (i + 1), DigitFrec.ScaleHeight - (peY * (MAX_DIGIT_WINDOW(i) * f(i)))), RGB(55, 55, 55), BF
    Next i
    '------------------------------------

End Function



Function frec(ByVal s As String)

    '------------------------------------
    Dim f(0 To 9) As Integer
    
    For i = 1 To Len(s)
    
        DI = Mid(s, i, 1)
    
        If DI = "0" Then f(0) = f(0) + 1
        If DI = "1" Then f(1) = f(1) + 1
        If DI = "2" Then f(2) = f(2) + 1
        If DI = "3" Then f(3) = f(3) + 1
        If DI = "4" Then f(4) = f(4) + 1
        If DI = "5" Then f(5) = f(5) + 1
        If DI = "6" Then f(6) = f(6) + 1
        If DI = "7" Then f(7) = f(7) + 1
        If DI = "8" Then f(8) = f(8) + 1
        If DI = "9" Then f(9) = f(9) + 1
    
    Next i
    
    
    For i = 0 To UBound(f)
        If f(i) > Maxim Then Maxim = f(i)
    Next i
    
    
    peX = (DigitFrec.ScaleWidth / (UBound(f) + 1))
    peY = (DigitFrec.ScaleHeight / Maxim)
    
    DigitFrec.Cls
    
    For i = 0 To UBound(f)
        'in functie de f maxima
        DigitFrec.Line (peX * i, DigitFrec.ScaleHeight)-(peX * (i + 1), DigitFrec.ScaleHeight - (peY * f(i))), RGB(55, 55, 55), BF

        'DigitFrec.Line (peX * i, DigitFrec.ScaleHeight / 2)-(peX * (i + 1), (peY * f(i))), RGB(55, 55, 55), BF
    Next i
    '------------------------------------

End Function


Private Sub Form_Load()

    Sec1.Text = Replace(Orig(0).Text, vbCrLf, "")
    Sec1.Text = Replace(Sec1.Text, vbCrLf, "")
    
    Call draw_scale(cicle)
    Window_Change
    Stepp_Change
    
    TR_PI_Click
    
End Sub

Private Sub Form_Terminate()
    stop_anim = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    stop_anim = True
End Sub


Private Sub Option2_Click(Index As Integer)

    If Option2(Index).Value = True Then
        a = Index
        Sec1.Text = Replace(Orig(a).Text, vbCrLf, "")
    End If

End Sub

Private Sub Stepp_Change()
    StepTxT.Caption = "Sliding window step: " & Stepp.Value
End Sub

Private Sub Stop_Run_Click()
    stop_anim = True
End Sub

Private Sub TR_PI_Click()

    Center_patt.Cls
    DigitFrec.Cls
    SumOfPI.Cls
    
    MatrixMaxVal = 0
    MatrixMinVal = 0
    
    Sec1.Text = Replace(Sec1.Text, vbCrLf, "")
    
    Call Fill_Transition_Matrix(M1, 9, 9, Sec1.Text)
    Call frec(Sec1.Text)
    Call SumPI2(Sec1.Text)
    Call P(Sec1.Text, 1)
    
    sTXT = DrowMatrix(9, 9, M1, "(P)", "Transition probabilities M:")
    sTXT_EXCEL = DrowMatrixForEXCEL(9, 9, M1, "(P)", "Copy/Paste to EXCEL:")
    
    WMatrix.Text = sTXT & sTXT_EXCEL
    
    For i = 0 To 9 '4
        For j = 0 To 9 '4
                If M1(i, j) > MatrixMaxVal Then MatrixMaxVal = M1(i, j)
                If M1(i, j) < MatrixMinVal Then MatrixMinVal = M1(i, j)
        Next j
    Next i
    
    Call DrowColorMatrix(9, 9, M1, Center_patt)

End Sub

Function Fill_Transition_Matrix(ByRef M() As String, ByVal cols As Integer, ByVal rows As Integer, ByVal s As String)

    For i = 0 To cols '4
    
        For j = 0 To rows '4
    
            M(i, j) = 0
    
        Next j
    
    Next i
    

    For i = 1 To Len(s)
    
        DI = Mid(s, i, 1)
    
        If DI = "0" Then n0 = n0 + 1
        If DI = "1" Then n1 = n1 + 1
        If DI = "2" Then n2 = n2 + 1
        If DI = "3" Then n3 = n3 + 1
        If DI = "4" Then n4 = n4 + 1
        If DI = "5" Then n5 = n5 + 1
        If DI = "6" Then n6 = n6 + 1
        If DI = "7" Then n7 = n7 + 1
        If DI = "8" Then n8 = n8 + 1
        If DI = "9" Then n9 = n9 + 1
    
    Next i
    
    
    For i = 1 To Len(s) - 1
    
            DI1 = Mid(s, i, 1)
            DI2 = Mid(s, i + 1, 1)
    
            If DI1 = "0" Then r = 0
            If DI1 = "1" Then r = 1
            If DI1 = "2" Then r = 2
            If DI1 = "3" Then r = 3
            If DI1 = "4" Then r = 4
            If DI1 = "5" Then r = 5
            If DI1 = "6" Then r = 6
            If DI1 = "7" Then r = 7
            If DI1 = "8" Then r = 8
            If DI1 = "9" Then r = 9
            
            If DI2 = "0" Then c = 0
            If DI2 = "1" Then c = 1
            If DI2 = "2" Then c = 2
            If DI2 = "3" Then c = 3
            If DI2 = "4" Then c = 4
            If DI2 = "5" Then c = 5
            If DI2 = "6" Then c = 6
            If DI2 = "7" Then c = 7
            If DI2 = "8" Then c = 8
            If DI2 = "9" Then c = 9
    
            M(r, c) = Val(M(r, c)) + 1
    
    Next i
    
    
    
    For i = 0 To cols
    
        For j = 0 To rows
    
            If n0 = 0 Then n0 = 1
            If n1 = 0 Then n1 = 1
            If n2 = 0 Then n2 = 1
            If n3 = 0 Then n3 = 1
            If n4 = 0 Then n4 = 1
            If n5 = 0 Then n5 = 1
            If n6 = 0 Then n6 = 1
            If n7 = 0 Then n7 = 1
            If n8 = 0 Then n8 = 1
            If n9 = 0 Then n9 = 1
            

            If i = 0 Then r = n0
            If i = 1 Then r = n1
            If i = 2 Then r = n3
            If i = 3 Then r = n3
            If i = 4 Then r = n4
            If i = 5 Then r = n5
            If i = 6 Then r = n6
            If i = 7 Then r = n7
            If i = 8 Then r = n8
            If i = 9 Then r = n9

            'M(i, j) = Round(Val(M(i, j)), 6) & "/" & r
            M(i, j) = Round(Val(M(i, j)) / r, 7)
    
        Next j
    
    Next i

End Function


Function DrowMatrix(ib, jb, ByVal M As Variant, ByVal model As String, ByVal msg As String) As String

    '------ Show Matrix in Text OBJ -------------------------------------------
    y = "|_____|_____|_____|_____|_____|_____|_____|_____|_____|_____|_____|"
    
    ct = ct & vbCrLf & "___________________________________________________________________"
    ct = ct & vbCrLf & "| " & model & " |  0  |  1  |  2  |  3  |  4  |  5  |  6  |  7  |  8  |  9  |"
    ct = ct & vbCrLf & y & vbCrLf
    
    For i = 0 To ib 'Rows
    
        For j = 0 To jb 'cols
        
            v = Round(M(i, j), 2)
        
            If Len(v) = 0 Then u = "|     "
            If Len(v) = 1 Then u = "|    "
            If Len(v) = 2 Then u = "|   "
            If Len(v) = 3 Then u = "|  "
            If Len(v) = 4 Then u = "| "
            If Len(v) = 5 Then u = "|"
            
            If j = jb Then o = "|" Else o = ""
            
            If j = 0 And i = 0 Then ct = ct & "|  0  "
            If j = 0 And i = 1 Then ct = ct & "|  1  "
            If j = 0 And i = 2 Then ct = ct & "|  2  "
            If j = 0 And i = 3 Then ct = ct & "|  3  "
            If j = 0 And i = 4 Then ct = ct & "|  4  "
            If j = 0 And i = 5 Then ct = ct & "|  5  "
            If j = 0 And i = 6 Then ct = ct & "|  6  "
            If j = 0 And i = 7 Then ct = ct & "|  7  "
            If j = 0 And i = 8 Then ct = ct & "|  8  "
            If j = 0 And i = 9 Then ct = ct & "|  9  "
            
            ct = ct & u & v & o
            
        Next j
    
    ct = ct & vbCrLf & y & vbCrLf
    
    Next i
    '--------------------------------------------------------------------------
    DrowMatrix = msg & " M[" & Val(jb) & "," & Val(ib) & "]" & vbCrLf & ct & vbCrLf & vbCrLf
    '--------------------------------------------------------------------------

End Function


Function DrowMatrixForEXCEL(ib, jb, ByVal M As Variant, ByVal model As String, ByVal msg As String) As String

    '------ Show Matrix in Text OBJ -------------------------------------------
    For i = 0 To ib 'Rows
    
        For j = 0 To jb 'cols
        
        v = Round(M(i, j), 2)
            
            If j = jb Then o = "" Else o = Chr(9)
            ct = ct & v & o
            
        Next j
    
    ct = ct & vbCrLf
    
    Next i
    '--------------------------------------------------------------------------
    DrowMatrixForEXCEL = msg & " M[" & Val(jb) & "," & Val(ib) & "]" & vbCrLf & ct & vbCrLf & vbCrLf
    '--------------------------------------------------------------------------

End Function


Function DrowColorMatrix(ib, jb, ByVal M As Variant, ByRef picOBJ As Object)

    '--------------------------------------------------------------------------
    picOBJ.Cls
    'Pic2.Cls
    
    Row = (picOBJ.ScaleWidth / (jb + 1))
    Col = (picOBJ.ScaleHeight / (ib + 1))
    
    Maxim = MatrixMaxVal
    Minim = MatrixMinVal
    
    
    If Maxim <> 0 Then Culoare1 = Int(255 / Maxim)
    If Minim <> 0 Then Culoare2 = Int(255 / Minim)
    
    
    For i = 0 To jb  'Rows
    
        For j = 0 To ib  'cols
        
            'h = M(i + 1, j + 1)
            h = M(j, i)
            
            If h > 0 Then r = Int(Culoare1 * h)
            If h < 0 Then g = Culoare2 * Abs(h)
            
            picOBJ.Line (Row * i, Col * j)-(Row * (i + 1), Col * (j + 1)), RGB(r, 55, 55), BF
            
            'If MatrixTrace(j + 1, i + 1) <> "" Then
            '    If PTB.Value = 1 Then Pic1.Line (Row * i, Col * j)-(Row * (i + 1), Col * (j + 1)), RGB(255, 255, 255), BF
            '    Pic2.Line (Row * i, Col * j)-(Row * (i + 1), Col * (j + 1)), RGB(200, 45, 45), BF
            '
            '    If PG.Value = 1 Then
            '        Pic2.Line (Row * i, 0)-(Row * i, Pic1.ScaleHeight), RGB(45, 45, 45), B
            '        Pic2.Line (0, Col * j)-(Pic1.ScaleWidth, Col * j), RGB(45, 45, 45), B
            '    End If
            'End If
            
        Next j
    
    Next i
    '--------------------------------------------------------------------------

End Function



Function draw_scale(ByVal k_stat As Integer)

    Dim zx, qx, zy, qy As Variant
    Dim sp As Variant
    Dim i As Integer
    
    Form1.Cls
    
    'X axis on Center_patt OBJ
    '-------------------------------------
    sp = Center_patt.ScaleWidth / 10
    
    For i = 0 To 9
    
        zx = Center_patt.Left + (sp * i)
        qx = zx
        zy = Center_patt.Top + Center_patt.ScaleHeight
        qy = Center_patt.Top - 10
        Form1.CurrentX = zx + (sp / 2)
        Form1.CurrentY = qy - 6
    
        
        If i = 0 Then Form1.Print "0"
        If i = 1 Then Form1.Print "1"
        If i = 2 Then Form1.Print "2"
        If i = 3 Then Form1.Print "3"
        If i = 4 Then Form1.Print "4"
        If i = 5 Then Form1.Print "5"
        If i = 6 Then Form1.Print "6"
        If i = 7 Then Form1.Print "7"
        If i = 8 Then Form1.Print "8"
        If i = 9 Then Form1.Print "9"
        
        Form1.Line (zx, zy)-(qx, qy), &H808080
    
    Next i
    '-------------------------------------
    
    'Y axis on Center_patt OBJ
    '-------------------------------------
    sp = Center_patt.ScaleHeight / 10
    
    For i = 0 To 9
    
        zx = Center_patt.Left - 6
        qx = Center_patt.Left
        zy = Center_patt.Top + (sp * i)
        qy = zy
        Form1.CurrentX = zx - 25
        Form1.CurrentY = qy + (sp / 2) - 6
        
        Form1.CurrentX = zx - 10
        
        If i = 0 Then Form1.Print "0"
        If i = 1 Then Form1.Print "1"
        If i = 2 Then Form1.Print "2"
        If i = 3 Then Form1.Print "3"
        If i = 4 Then Form1.Print "4"
        If i = 5 Then Form1.Print "5"
        If i = 6 Then Form1.Print "6"
        If i = 7 Then Form1.Print "7"
        If i = 8 Then Form1.Print "8"
        If i = 9 Then Form1.Print "9"
    
        Form1.Line (zx, zy)-(qx, qy), &H808080
    
    Next i
    '-------------------------------------
    
    'X axis on DigitFrec OBJ
    '-------------------------------------
    sp = DigitFrec.ScaleWidth / 10
    
    For i = 0 To 9
    
        zx = DigitFrec.Left + (sp * i)
        qx = zx
        zy = DigitFrec.Top + DigitFrec.ScaleHeight
        qy = DigitFrec.Top + DigitFrec.ScaleHeight + 6
        Form1.CurrentX = zx + (sp / 2)
        Form1.CurrentY = qy
    
        
        If i = 0 Then Form1.Print "0"
        If i = 1 Then Form1.Print "1"
        If i = 2 Then Form1.Print "2"
        If i = 3 Then Form1.Print "3"
        If i = 4 Then Form1.Print "4"
        If i = 5 Then Form1.Print "5"
        If i = 6 Then Form1.Print "6"
        If i = 7 Then Form1.Print "7"
        If i = 8 Then Form1.Print "8"
        If i = 9 Then Form1.Print "9"
        
        Form1.Line (zx, zy)-(qx, qy), &H808080
    
    Next i
    '-------------------------------------
    
    
    'X axis on DigitFrec OBJ
    '-------------------------------------
    sp = MaxFWindow.ScaleWidth / 10
    
    For i = 0 To 9
    
        zx = DigitFrec.Left + (sp * i)
        qx = zx
        zy = MaxFWindow.Top + MaxFWindow.ScaleHeight
        qy = MaxFWindow.Top + MaxFWindow.ScaleHeight + 6
        Form1.CurrentX = zx + (sp / 2)
        Form1.CurrentY = qy
    
        
        If i = 0 Then Form1.Print "0"
        If i = 1 Then Form1.Print "1"
        If i = 2 Then Form1.Print "2"
        If i = 3 Then Form1.Print "3"
        If i = 4 Then Form1.Print "4"
        If i = 5 Then Form1.Print "5"
        If i = 6 Then Form1.Print "6"
        If i = 7 Then Form1.Print "7"
        If i = 8 Then Form1.Print "8"
        If i = 9 Then Form1.Print "9"
        
        Form1.Line (zx, zy)-(qx, qy), &H808080
    
    Next i
    '-------------------------------------

End Function


Private Sub Window_Change()
    WindowTXT.Caption = "Sliding window length: " & Window.Value
End Sub



Function RulePI(ByVal s As String, ByVal rule As Integer) As String

    Dim z() As String
    
    '###############################################################################################
    'x-tuple
    '------------------------------------
    digits = x_tuple.Value
    
    a = 1
    
    'Debug.Print "-------------------"
    
    For i = 1 To Len(s)
        DI = Val(Mid(s, i, digits))
        DI2 = Val(Mid(s, i + digits, digits))
    
        'If i < 30 Then Debug.Print DI & "-" & DI2
        
        If rule = 0 Then
            If (Val(DI) > Val(DI2)) Then a = a + 1 Else a = a - 1
        End If
    
        If rule = 1 Then
            If (Val(DI) < Val(DI2)) Then a = a + 1 Else a = a - 1
        End If
    
        c = c & "," & a '& "[" & DI & "]"
        'c1 = c1 & Chr(9) & a
        c1 = c1 & ", " & a
        
    Next i
    
    RulePI = c1
    '------------------------------------
    z() = Split(c, ",")
    
    If OSignal.Value = 0 Then MaxFWindow.Cls
    
    For i = 1 To UBound(z)
        If Abs(Maxz) < Abs(z(i)) Then Maxz = Abs(z(i))
    Next i
    
    xxx = MaxFWindow.ScaleWidth / UBound(z)
    yyy = MaxFWindow.ScaleHeight / Maxz
    
    For i = 1 To UBound(z)
        'MaxFWindow.Line (xxx * i, MaxFWindow.ScaleHeight - (yyy * Abs(z(i))))-(xxx * i, MaxFWindow.ScaleHeight - (yyy * Abs(z(i)))), vbRed, BF
        MaxFWindow.Line (xxx * i, yyy * Abs(z(i)))-(xxx * i, yyy * Abs(z(i))), vbRed, BF
    Next i
    
    '###############################################################################################

End Function

Private Sub x_tuple_Change()

    digits = x_tuple.Value
    s = Sec1.Text
    tmp = ""
    
    For i = 1 To 20
        DI = Val(Mid(s, i, digits))
        DI2 = Val(Mid(s, i + digits, digits))
        
        tmp = tmp & "[" & DI & "|" & DI2 & "], "
        
    Next i
    
    Arata_tuple.Caption = Mid(tmp, 1, 30) & "..."
    
End Sub
