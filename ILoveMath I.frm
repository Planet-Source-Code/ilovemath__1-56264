VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form oILoveMath 
   BackColor       =   &H000040C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Prototipe II IloveMath"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6390
   Icon            =   "ILoveMath I.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   6390
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      BackColor       =   &H000040C0&
      Caption         =   "Time"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3840
      TabIndex        =   11
      Top             =   4080
      Width           =   2415
      Begin MSComctlLib.ProgressBar progWaktu 
         Height          =   135
         Left            =   720
         TabIndex        =   12
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   0
         Max             =   5000
         Scrolling       =   1
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H000040C0&
      Caption         =   "Score"
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   1920
      TabIndex        =   9
      Top             =   4080
      Width           =   1815
      Begin VB.Label lblScore 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "9 X 1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H000040C0&
      Caption         =   "Question"
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   120
      TabIndex        =   7
      Top             =   4080
      Width           =   1695
      Begin VB.Label lblPerhitungan 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "9 X 1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Left            =   0
         TabIndex        =   8
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Timer tmrChecker 
      Interval        =   1
      Left            =   1080
      Top             =   5760
   End
   Begin VB.Timer tmrWaktu 
      Interval        =   5
      Left            =   720
      Top             =   5760
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000040C0&
      Caption         =   "Power"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3840
      TabIndex        =   2
      Top             =   4680
      Width           =   2415
      Begin MSComctlLib.ProgressBar progNyawa 
         Height          =   135
         Left            =   720
         TabIndex        =   6
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   0
         Max             =   90
         Scrolling       =   1
      End
   End
   Begin VB.Timer tmrChar 
      Interval        =   500
      Left            =   360
      Top             =   5760
   End
   Begin VB.PictureBox picKanvas 
      BackColor       =   &H00404040&
      Height          =   3855
      Left            =   120
      Picture         =   "ILoveMath I.frx":0442
      ScaleHeight     =   3795
      ScaleWidth      =   6075
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      Begin VB.Image imgBenda 
         Height          =   480
         Index           =   0
         Left            =   1200
         Picture         =   "ILoveMath I.frx":47114
         Top             =   360
         Width           =   480
      End
      Begin VB.Image imgBenda 
         Height          =   480
         Index           =   1
         Left            =   3960
         Picture         =   "ILoveMath I.frx":47556
         Top             =   0
         Width           =   480
      End
      Begin VB.Image imgBenda 
         Height          =   480
         Index           =   2
         Left            =   2640
         Picture         =   "ILoveMath I.frx":47998
         Top             =   2160
         Width           =   480
      End
      Begin VB.Image imgBenda 
         Height          =   480
         Index           =   3
         Left            =   5400
         Picture         =   "ILoveMath I.frx":47DDA
         Top             =   3000
         Width           =   480
      End
      Begin VB.Image imgBenda 
         Height          =   480
         Index           =   4
         Left            =   240
         Picture         =   "ILoveMath I.frx":4821C
         Top             =   1680
         Width           =   480
      End
      Begin VB.Image imgBenda 
         Height          =   480
         Index           =   5
         Left            =   4920
         Picture         =   "ILoveMath I.frx":4865E
         Top             =   1320
         Width           =   480
      End
      Begin VB.Image ImgPesawat 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   2760
         Picture         =   "ILoveMath I.frx":48AA0
         Top             =   3120
         Width           =   480
      End
      Begin VB.Label lblHasil 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Index           =   3
         Left            =   5490
         TabIndex        =   5
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label lblHasil 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Index           =   2
         Left            =   4060
         TabIndex        =   4
         Top             =   2730
         Width           =   375
      End
      Begin VB.Label lblHasil 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Index           =   1
         Left            =   1320
         TabIndex        =   3
         Top             =   2730
         Width           =   375
      End
      Begin VB.Image imgChecker4 
         Height          =   495
         Left            =   4710
         Top             =   705
         Width           =   495
      End
      Begin VB.Image imgChecker3 
         Height          =   495
         Left            =   3510
         Top             =   1815
         Width           =   495
      End
      Begin VB.Image imgChecker2 
         Height          =   495
         Left            =   1755
         Top             =   1830
         Width           =   495
      End
      Begin VB.Image imgChecker1 
         Height          =   495
         Left            =   720
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblHasil 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Index           =   0
         Left            =   160
         TabIndex        =   1
         Top             =   1320
         Width           =   375
      End
   End
   Begin VB.Timer tmrArah 
      Interval        =   1
      Left            =   0
      Top             =   5760
   End
End
Attribute VB_Name = "oILoveMath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oColl As Integer

Private Sub Form_Load()
    progNyawa.Value = 90
    progWaktu.Value = 5000
    lblScore.Caption = 0
    lblPerhitungan = oPerhitungan(lblHasil(0), lblHasil(1), lblHasil(2), lblHasil(3))
End Sub

Private Sub tmrChar_Timer() 'Anim the alternative answer
    For Q = 0 To 3
        If lblHasil(Q).ForeColor = &H404040 Then
            lblHasil(Q).ForeColor = &HFFFFFF
        ElseIf lblHasil(Q).ForeColor = &HFFFFFF Then
            lblHasil(Q).ForeColor = &H404040
        End If
    Next Q
End Sub

Private Sub tmrChecker_Timer() 'Check position
    oCheckPos
End Sub

Private Sub tmrArah_Timer()  'Move the starship
    If GetAsyncKeyState(vbKeyUp) Then
        oTombol = 1
        oJalankan
    ElseIf GetAsyncKeyState(vbKeyDown) Then
        oTombol = 2
        oJalankan
    ElseIf GetAsyncKeyState(vbKeyRight) Then
        oTombol = 3
        oJalankan
    ElseIf GetAsyncKeyState(vbKeyLeft) Then
        oTombol = 4
        oJalankan
    End If
    
    oAwan
End Sub

Private Sub oJalankan() 'Move the starship
   Select Case oTombol
        Case 1
            ImgPesawat.Top = ImgPesawat.Top - 30
        Case 2
            ImgPesawat.Top = ImgPesawat.Top + 30
        Case 3
            ImgPesawat.Left = ImgPesawat.Left + 30
        Case 4
            ImgPesawat.Left = ImgPesawat.Left - 30
   End Select
   
   If oTombol = 1 Or oTombol = 2 Then
        If ImgPesawat.Top < 0 Then
            ImgPesawat.Top = 10
        ElseIf ImgPesawat.Top > 3290 Then
            ImgPesawat.Top = 3280
        End If
   ElseIf oTombol = 3 Or oTombol = 4 Then
        If ImgPesawat.Left < 0 Then
            ImgPesawat.Left = 10
        ElseIf ImgPesawat.Left > 5590 Then
            ImgPesawat.Left = 5580
        End If
   End If
End Sub

Private Sub oAwan()  'Cloud
    Randomize
    For i = 0 To imgBenda.UBound
        imgBenda(i).Top = imgBenda(i).Top + 10
        
        If imgBenda(i).Top > 3600 Then
            imgBenda(i).Top = -30
            imgBenda(i).Left = CInt(Rnd * 5000)
        End If
    Next i
End Sub

Private Sub oCheckPos()
    If ImgPesawat.Top > 600 And ImgPesawat.Top < 700 Then
       If ImgPesawat.Left > 700 And ImgPesawat.Left < 900 Then
            oColl = 1
            oCekJWB
       End If
    End If
    
    If ImgPesawat.Top > 1825 And ImgPesawat.Top < 1900 Then
        If ImgPesawat.Left > 1700 And ImgPesawat.Left < 1800 Then
            oColl = 2
            oCekJWB
        End If
    End If
    
    If ImgPesawat.Top > 1820 And ImgPesawat.Top < 1920 Then
        If ImgPesawat.Left > 3510 And ImgPesawat.Left < 3610 Then
            oColl = 3
            oCekJWB
        End If
    End If
    
    If ImgPesawat.Top > 710 And ImgPesawat.Top < 810 Then
        If ImgPesawat.Left > 4710 And ImgPesawat.Left < 4810 Then
            oColl = 4
            oCekJWB
        End If
    End If
End Sub

Private Sub oCekJWB() 'Check the answer
   If oColl = oPos Then
        MsgBox "yeah right !", vbInformation, "Peringatan"
        lblPerhitungan = oPerhitungan(lblHasil(0), lblHasil(1), lblHasil(2), lblHasil(3))
        lblScore = Val(lblScore) + 1
        progWaktu.Value = 5000
        
        ImgPesawat.Top = 3120
        ImgPesawat.Left = 2760
    Else
        MsgBox "Ops it's not correct answer !", vbCritical, "Peringatan"
        progNyawa.Value = progNyawa.Value - 30
        
        oComment
        
        ImgPesawat.Top = 3120
        ImgPesawat.Left = 2760
    End If
End Sub

Private Sub tmrWaktu_Timer() 'Time limit
    progWaktu.Value = progWaktu.Value - 1
    
    If progWaktu.Value = 0 Then
        MsgBox "Time over", vbCritical, "Peringatan"
        
        lblPerhitungan = oPerhitungan(lblHasil(0), lblHasil(1), lblHasil(2), lblHasil(3))
        progNyawa.Value = progNyawa.Value - 30
        
        oComment
        
        ImgPesawat.Top = 3120
        ImgPesawat.Left = 2760
        progWaktu.Value = 5000
    End If
End Sub

Private Sub oComment()
    If progNyawa.Value = 0 Then
            MsgBox "GAME OVER!", vbCritical, "Peringatan"
            progNyawa.Value = 90
            
            lblPerhitungan = oPerhitungan(lblHasil(0), lblHasil(1), lblHasil(2), lblHasil(3))
            lblScore.Caption = 0
    End If
End Sub
