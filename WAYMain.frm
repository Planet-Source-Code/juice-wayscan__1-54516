VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Main 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "WAY"
   ClientHeight    =   3855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4695
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00800000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox RemColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Replace excluded colors.."
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   2235
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Shows the percentage of scan completion, but slows down the scan slightly"
      Top             =   3450
      Width           =   2415
   End
   Begin VB.CheckBox ShowPercentage 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Show scan percentage"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Shows the percentage of scan completion, but slows down the scan slightly"
      Top             =   3450
      Width           =   2055
   End
   Begin VB.CheckBox GenerateLogFile 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Generate Log File"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Generates a log file to display matrix information, but slows down the scan"
      Top             =   3195
      Width           =   1650
   End
   Begin VB.CheckBox UseExclusions 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Honor exclusions during scan"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   2235
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Makes scanning much more accurate, but slows down the scan"
      Top             =   3195
      Width           =   2400
   End
   Begin VB.PictureBox IMG 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2190
      Left            =   120
      ScaleHeight     =   144
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   296
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   960
      Width           =   4470
   End
   Begin VB.Timer DelSel 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   540
      Top             =   2775
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   45
      Top             =   2745
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label ClearExclusions 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Clear Exclusions"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   2640
      TabIndex        =   10
      Top             =   645
      Width           =   1335
   End
   Begin VB.Label SetExclusions 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Set Exclusions"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      Top             =   645
      Width           =   1335
   End
   Begin VB.Label ScanFile 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Scan Image"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   645
      Width           =   975
   End
   Begin VB.Label CurrentFileSize 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "(0 bytes)"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   3345
      TabIndex        =   4
      Top             =   345
      Width           =   660
   End
   Begin VB.Label xBtn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "x"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4455
      TabIndex        =   3
      ToolTipText     =   "Exit"
      Top             =   0
      Width           =   255
   End
   Begin VB.Label CurrentFile 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " None selected"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   330
      Width           =   2055
   End
   Begin VB.Label SelectFile 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Select File"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   330
      Width           =   975
   End
   Begin VB.Shape FormBorder 
      BorderColor     =   &H00800000&
      Height          =   3645
      Left            =   0
      Top             =   210
      Width           =   4695
   End
   Begin VB.Label ToolBar 
      BackColor       =   &H00800000&
      Caption         =   "  WAYScan"
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Dim IMGFile As String
Dim MouseD As Boolean
Dim MouseX As Single
Dim MouseY As Single
Dim Matrix(65535) As String
Dim MatrixPlots As Long
Dim RepColor As Long
Dim Exclude As Boolean
Dim HighlightOff As OLE_COLOR

Private Sub ClearExclusions_Click()
    MatrixPlots = 0
        If IMGFile <> "" Then
            IMG.Cls
            IMG.PaintPicture LoadPicture(IMGFile), 0, 0, IMG.ScaleWidth, IMG.ScaleHeight
        End If
    Exclude = False
    ToolBar.Caption = "  WAYScan"
End Sub

Private Sub ClearExclusions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Sel ClearExclusions
End Sub

Private Sub DelSel_Timer()
    KillSel
    DelSel.Enabled = False
End Sub

Private Sub Form_Load()
    HighlightOff = &HE0E0E0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    KillSel
End Sub

Private Sub KillSel()
    SelectFile.BackColor = HighlightOff
    ScanFile.BackColor = HighlightOff
    SetExclusions.BackColor = HighlightOff
    ClearExclusions.BackColor = HighlightOff
End Sub

Private Sub NoSel(Target As Label)
    Target.BackColor = HighlightOff
    DelSel.Enabled = False
End Sub

Private Sub Sel(Target As Label)
    Target.BackColor = vbWhite
    DelSel.Enabled = True
End Sub

Private Sub IMG_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Ret As Long
    If MatrixPlots >= 65535 Then
        MsgBox "You've selected the maximum number of colors to exclude.", vbInformation
    Else
        If Exclude Then
            Ret = GetPixel(IMG.hdc, X, Y)
                For i = 0 To MatrixPlots
                    If CStr(Ret) = Matrix(i) Then Exit Sub
                Next
            Matrix(MatrixPlots) = Ret
            MatrixPlots = MatrixPlots + 1
            ToolBar.Caption = "  WAYScan (" & MatrixPlots & " colors excluded)"
            IMG.PSet (X, Y), vbRed
            UseExclusions.Enabled = True
        End If
    End If
End Sub

Private Sub IMG_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Ret As Long
KillSel
    If Exclude And Shift = 1 Then
        If MatrixPlots >= 65535 Then
            MsgBox "You've selected the maximum number of colors to exclude.", vbInformation
        Else
            Ret = GetPixel(IMG.hdc, X, Y)
                For i = 0 To MatrixPlots
                    If CStr(Ret) = Matrix(i) Then Exit Sub
                Next
            Matrix(MatrixPlots) = Ret
            MatrixPlots = MatrixPlots + 1
            ToolBar.Caption = "  WAYScan (" & MatrixPlots & " colors excluded)"
            IMG.PSet (X, Y), vbRed
        End If
    End If
End Sub

Private Sub RemColor_Click()
    If RemColor.Value = 1 Then
        CD.DialogTitle = "Select replacement color"
        CD.ShowColor
        RepColor = CD.Color
    End If
End Sub

Private Sub ScanFile_Click()
Dim Ret As Long
Dim Skin As String
Dim BIntegrity As Long
Dim WIntegrity As Long
Dim Accuracy As Long
Dim ExcludePixel As Boolean
Dim RepPixel As Boolean
Dim DivideBy As Long
Dim PercentDone As Long
Dim Completed As Integer
    If IMGFile = "" Then
        MsgBox "No image is selected", vbInformation
        Exit Sub
    End If
    GenerateLogFile.Enabled = False
    UseExclusions.Enabled = False
    ShowPercentage.Enabled = False
    RemColor.Enabled = False
    DivideBy = IMG.ScaleHeight * IMG.ScaleWidth
    DivideBy = DivideBy / 100
    ToolBar.Caption = "  WAYScan - Analyzing image.."
    DoEvents
        If GenerateLogFile.Value = 1 Then
            Open App.Path & "\WAYScan.log" For Output As #1
        End If
            For a = 0 To IMG.ScaleHeight
                For b = 0 To IMG.ScaleWidth
                    If ShowPercentage.Value = 1 Then
                        DoEvents
                        Completed = CInt(PercentDone / DivideBy)
                        If Completed > 100 Then Completed = 100
                        ToolBar.Caption = "  WAYScan - Analyzing image (" & Completed & "%)"
                    End If
                    ExcludePixel = False
                    RepPixel = False
                    Ret = GetPixel(IMG.hdc, b, a)
                        If RemColor.Value = 1 Then
                            For c = 0 To MatrixPlots
                                If CStr(Ret) = Matrix(c) Then RepPixel = True
                            Next
                        End If
                        If UseExclusions.Value = 1 Then
                            For c = 0 To MatrixPlots
                                If CStr(Ret) = Matrix(c) Then ExcludePixel = True
                            Next
                        End If
                    If Ret = 16777215 And GenerateLogFile.Value = 1 Then
                        Print #1, "Matrix(" & b & "," & a & ") = " & Ret & " (Null space)"
                    Else
                        If ExcludePixel = False Then
                            If Ret > 2000000 And Ret < 5000000 Then
                                BIntegrity = BIntegrity + 1
                            ElseIf Ret > 6000000 And Ret < 16000000 Then
                                WIntegrity = WIntegrity + 1
                            End If
                                If GenerateLogFile.Value = 1 Then
                                    Print #1, "Matrix(" & b & "," & a & ") = " & Ret
                                End If
                        Else
                            If GenerateLogFile.Value = 1 Then
                                If RepPixel Then
                                    Dim sRet As Long
                                    sRet = SetPixel(IMG.hdc, b, a, RepColor)
                                    Print #1, "Matrix(" & b & "," & a & ") = " & Ret & " (Excluded by user and color replaced with " & RepColor & " - Return: " & sRet & ")"
                                Else
                                    Print #1, "Matrix(" & b & "," & a & ") = " & Ret & " (Excluded by user)"
                                End If
                            End If
                        End If
                    End If
                PercentDone = PercentDone + 1
                Next
            Next
        ToolBar.Caption = "  WAYScan"
    If BIntegrity < 500 And WIntegrity < 500 Then
        Accuracy = 100
        Skin = "neither black or white."
    Else
        If BIntegrity > WIntegrity Then
            Accuracy = BIntegrity - WIntegrity
            Accuracy = Accuracy / 100
            If Accuracy > 100 Then Accuracy = 100
               Accuracy = 100 - Accuracy
            If Accuracy = 0 Then Accuracy = 100
            If Accuracy < 50 Then Accuracy = 100 - Accuracy
            Skin = "black (estimating about " & Accuracy & "% accuracy)"
        Else
            Accuracy = WIntegrity - BIntegrity
            Accuracy = Accuracy / 100
            If Accuracy > 100 Then Accuracy = 100
               Accuracy = 100 - Accuracy
            If Accuracy = 0 Then Accuracy = 100
            If Accuracy < 50 Then Accuracy = 100 - Accuracy
            Skin = "white (estimating about " & Accuracy & "% accuracy)"
        End If
    End If
        If GenerateLogFile.Value = 1 Then
            Print #1, "WAYScan has determined that this individual is most likely " & Skin & "."
            Close #1
            Reset
        End If
IMG.Refresh
MsgBox "WAYScan has determined that this individual is most likely " & Skin & ".", vbInformation
GenerateLogFile.Enabled = True
UseExclusions.Enabled = True
ShowPercentage.Enabled = True
RemColor.Enabled = True
End Sub

Private Sub ScanFile_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Sel ScanFile
End Sub

Private Sub SelectFile_Click()
    NoSel SelectFile
    CD.DialogTitle = "Select an image"
    CD.FileName = ""
    CD.Filter = "Images|*.jpg;*.bmp;*.tga;*.png;*.gif"
    CD.ShowOpen
        If CD.FileName <> "" Then
            IMGFile = CD.FileName
            CurrentFile.Caption = " " & CD.FileTitle
            CurrentFileSize.Caption = "(" & FileLen(IMGFile) & " bytes)"
            IMG.Cls
            IMG.PaintPicture LoadPicture(IMGFile), 0, 0, IMG.ScaleWidth, IMG.ScaleHeight
                If Exclude Then
                    Exclude = False
                    MatrixPlots = 0
                End If
        End If
End Sub

Private Sub SelectFile_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Sel SelectFile
End Sub

Private Sub SetExclusions_Click()
    Exclude = True
    MsgBox "Select the colors to exclude from processing." & vbCrLf & _
           "Hold SHIFT to select without clicking.", vbInformation
End Sub

Private Sub SetExclusions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Sel SetExclusions
End Sub

Private Sub ToolBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        MouseD = True
        MouseX = X
        MouseY = Y
    End If
End Sub

Private Sub ToolBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If MouseD Then
        Me.Left = Me.Left - (MouseX - X)
        Me.Top = Me.Top - (MouseY - Y)
    End If
End Sub

Private Sub ToolBar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        MouseD = False
    End If
End Sub

Sub UpdateSize()
    CurrentFileSize.Caption = "(" & FileLen(IMGFile) & " bytes)"
End Sub

Private Sub xBtn_Click()
    Unload Me
    End
End Sub
