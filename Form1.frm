VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   8145
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13650
   LinkTopic       =   "Form1"
   ScaleHeight     =   543
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   910
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Pic 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2550
      Index           =   7
      Left            =   3900
      ScaleHeight     =   170
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   196
      TabIndex        =   14
      Top             =   195
      Width           =   2940
      Begin VB.Label lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Index           =   7
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   270
      End
   End
   Begin VB.PictureBox Pic 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2745
      Index           =   6
      Left            =   6825
      ScaleHeight     =   183
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   183
      TabIndex        =   12
      Top             =   1950
      Width           =   2745
      Begin VB.Label lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Index           =   6
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   270
      End
   End
   Begin VB.PictureBox Pic 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   1770
      Index           =   5
      Left            =   6240
      ScaleHeight     =   118
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   300
      TabIndex        =   10
      Top             =   0
      Width           =   4500
      Begin VB.Label lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Index           =   5
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   270
      End
   End
   Begin VB.PictureBox Pic 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   4110
      Index           =   4
      Left            =   1950
      ScaleHeight     =   274
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   378
      TabIndex        =   4
      Top             =   3315
      Width           =   5670
      Begin VB.Label lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Index           =   4
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   270
      End
   End
   Begin VB.PictureBox Pic 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1800
      Index           =   3
      Left            =   195
      ScaleHeight     =   120
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   170
      TabIndex        =   3
      Top             =   195
      Width           =   2550
      Begin VB.Label lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Index           =   3
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   270
      End
   End
   Begin VB.PictureBox Pic 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   4110
      Index           =   2
      Left            =   195
      ScaleHeight     =   274
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   118
      TabIndex        =   2
      Top             =   3315
      Width           =   1770
      Begin VB.Label lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Index           =   2
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   270
      End
   End
   Begin VB.PictureBox Pic 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   795
      Index           =   1
      Left            =   195
      ScaleHeight     =   53
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   235
      TabIndex        =   1
      Top             =   2145
      Width           =   3525
      Begin VB.Label lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Index           =   1
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   270
      End
   End
   Begin VB.PictureBox Pic 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1800
      Index           =   0
      Left            =   4095
      ScaleHeight     =   120
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   120
      TabIndex        =   0
      Top             =   3705
      Width           =   1800
      Begin VB.Label lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Index           =   0
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   240
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Enum Justifications
    arrangeLeft = 1
    arrangeRight = 2
    arrangeTop = 3
    arrangeBottom = 4
End Enum

Dim imgItem()
Dim imgX()
Dim imgY()
Dim imgSize()
Dim GridSize As Long 'in Pixels
Dim GridData() As Boolean
Dim MaxGridX As Long
Dim MaxGridY As Long
Dim ShowGrid As Boolean
Dim Justification As Long


'// *************************************************************************************************
Private Sub Form_Load()
    ShowGrid = True
    Justification = 1
    Call Init
End Sub

Private Sub Init()
    GridSize = 14
    Call DrawGrid(GridSize, Me, ShowGrid) 'in Pixels
    Call ArrangeAll(Justification)
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.Cls
    Justification = Justification + 1
    If Justification > 4 Then Justification = 1
    ShowGrid = (ShowGrid = False)
    Call Init
End Sub

Private Sub DrawGrid(ByVal pSize As Long, ByRef Frm As Form, ByVal ShowGridLines As Boolean)
For X = 0 To Frm.ScaleWidth - (pSize * 1) Step pSize
   For Y = 0 To Frm.ScaleHeight - (pSize * 1) Step pSize
       If ShowGridLines Then Me.Line (X, Y)-(X + pSize, Y + pSize), QBColor(0), B
   Next
Next
MaxGridX = Int(X / GridSize)
MaxGridY = Int(Y / GridSize)
Me.Caption = MaxGridX & " x " & MaxGridY
ReDim GridData(MaxGridX, MaxGridY)
End Sub

Private Sub ArrangeWindows()

'// Get Width, Height, and Size
Cnt = -1
For Each Control In Controls
    If TypeOf Control Is PictureBox Then
        Cnt = Cnt + 1
        ReDim Preserve imgItem(Cnt)
        imgItem(Cnt) = Cnt
        
        ReDim Preserve imgX(Cnt)
        imgX(Cnt) = Pic(Cnt).ScaleWidth
        
        ReDim Preserve imgY(Cnt)
        imgY(Cnt) = Pic(Cnt).ScaleHeight
        
        ReDim Preserve imgSize(Cnt)
        imgSize(Cnt) = imgX(Cnt) * imgY(Cnt)
    End If
Next
ShellSort imgSize, , True

'// Arrange them now based on Size
For i = 0 To Cnt
    gx = Int(imgX(i) / GridSize)
    If (imgX(i) Mod GridSize) > 0 Then gx = gx + 1
    
    gy = Int(imgY(i) / GridSize)
    If (imgY(i) Mod GridSize) > 0 Then gy = gy + 1
    
    foundit = False
    abort = False
    
    For Y = 0 To MaxGridY
        For X = 0 To MaxGridX
            IsClear = GridIsClear(X, Y, gx, gy)
            If IsClear Then
                x2 = X + gx
                y2 = Y + gy
                foundit = True
                Call SetGrid(X, Y, x2, y2)
                Pic(imgItem(i)).Move X * GridSize, Y * GridSize
                Me.lab(imgItem(i)).Caption = i + 1
                Exit For
            End If
        Next
        If foundit Then Exit For
    Next
    If Not foundit Then
        '// Exception code (What to do if there's no space left for this object
        Pic(imgItem(i)).Move (MaxGridX * GridSize) - (gx * GridSize), (MaxGridY * GridSize) - (gy * GridSize)
        Me.lab(imgItem(i)).Caption = i + 1
    End If
Next
End Sub

Private Function ArrangeAll(Optional ByVal Justify As Justifications)
    
    LockWindowUpdate Me.hWnd
    If Justify = 0 Then Justify = 3
    
    '// Get Width, Height, and Size
    Cnt = -1
    For Each Control In Controls
        If TypeOf Control Is PictureBox Then
            Cnt = Cnt + 1
            ReDim Preserve imgItem(Cnt)
            imgItem(Cnt) = Cnt
            
            ReDim Preserve imgX(Cnt)
            imgX(Cnt) = Pic(Cnt).ScaleWidth
            
            ReDim Preserve imgY(Cnt)
            imgY(Cnt) = Pic(Cnt).ScaleHeight
            
            ReDim Preserve imgSize(Cnt)
            imgSize(Cnt) = imgX(Cnt) * imgY(Cnt)
        End If
    Next
    ShellSort imgSize, , True
    
    
    '// Arrange based on Size
    For i = 0 To Cnt
        gx = Int(imgX(i) / GridSize)
        If (imgX(i) Mod GridSize) > 0 Then gx = gx + 1
        gy = Int(imgY(i) / GridSize)
        If (imgY(i) Mod GridSize) > 0 Then gy = gy + 1
        foundit = False
        abort = False
        
        Select Case Justify
        Case arrangeTop
            For Y = 0 To MaxGridY
                For X = 0 To MaxGridX
                    IsClear = GridIsClear(X, Y, gx, gy)
                    If IsClear Then
                        foundit = True
                        Call SetGrid(X, Y, X + gx, Y + gy)
                        Pic(imgItem(i)).Move X * GridSize, Y * GridSize
                        Me.lab(imgItem(i)).Caption = i + 1
                        Exit For
                    End If
                Next
                If foundit Then Exit For
            Next
        Case arrangeBottom
            For Y = MaxGridY To 0 Step -1
                For X = 0 To MaxGridX
                    IsClear = GridIsClear(X, Y, gx, gy)
                    If IsClear Then
                        foundit = True
                        Call SetGrid(X, Y, X + gx, Y + gy)
                        Pic(imgItem(i)).Move X * GridSize, Y * GridSize
                        Me.lab(imgItem(i)).Caption = i + 1
                        Exit For
                    End If
                Next
                If foundit Then Exit For
            Next
        Case arrangeLeft
            For X = 0 To MaxGridX
                For Y = 0 To MaxGridY
                    IsClear = GridIsClear(X, Y, gx, gy)
                    If IsClear Then
                        foundit = True
                        Call SetGrid(X, Y, X + gx, Y + gy)
                        Pic(imgItem(i)).Move X * GridSize, Y * GridSize
                        Me.lab(imgItem(i)).Caption = i + 1
                        Exit For
                    End If
                Next
                If foundit Then Exit For
            Next
        Case arrangeRight
            For X = MaxGridX To 0 Step -1
                For Y = 0 To MaxGridY
                    IsClear = GridIsClear(X, Y, gx, gy)
                    If IsClear Then
                        foundit = True
                        Call SetGrid(X, Y, X + gx, Y + gy)
                        Pic(imgItem(i)).Move X * GridSize, Y * GridSize
                        Me.lab(imgItem(i)).Caption = i + 1
                        Exit For
                    End If
                Next
                If foundit Then Exit For
            Next
        End Select
        If Not foundit Then '// Exception code
            Pic(imgItem(i)).Move (MaxGridX * GridSize) - (gx * GridSize), (MaxGridY * GridSize) - (gy * GridSize)
            Me.lab(imgItem(i)).Caption = i + 1
        End If
        foundit = False
    Next
LockWindowUpdate 0
End Function

Private Sub SetGrid(X, Y, x2, y2)
For xx = X To (x2 - 1)
    For yy = Y To (y2 - 1)
        GridData(xx, yy) = True
        xxx = xx * GridSize
        yyy = yy * GridSize
        Me.Line (xxx, yyy)-(xxx + GridSize, yyy + GridSize), QBColor(2), BF
    Next yy
Next xx
End Sub

Private Function GridIsClear(X, Y, gx, gy) As Boolean
GridIsClear = False
abort = False
For xx = X To (X + gx)
    If (X + gx) > MaxGridX Then
        abort = True
        Exit For
    End If
    For yy = Y To (Y + gy)
        If (Y + gy) > MaxGridY Then
            abort = True
            Exit For
        End If
        If GridData(xx, yy) = True Then
            abort = True
            Exit For
        End If
    Next yy
    If abort = True Then Exit For
Next xx
If Not abort Then GridIsClear = True
End Function

Sub ShellSort(arr As Variant, Optional lastEl As Variant, Optional descending As Boolean)
    Dim value As Variant
    Dim index As Long, index2 As Long
    Dim firstEl As Long
    Dim distance As Long
    Dim numEls As Long

    ' account for optional arguments
    If IsMissing(lastEl) Then lastEl = UBound(arr)
    firstEl = LBound(arr)
    
    numEls = lastEl - firstEl + 1
    ' find the best value for distance
    Do
        distance = distance * 3 + 1
    Loop Until distance > numEls

    Do
        distance = distance \ 3
        For index = distance + firstEl To lastEl
            value1 = arr(index)
            value2 = imgItem(index)
            value3 = imgX(index)
            value4 = imgY(index)
            
            index2 = index
            Do While (arr(index2 - distance) > value1) Xor descending
                arr(index2) = arr(index2 - distance)
                
                imgItem(index2) = imgItem(index2 - distance)
                imgX(index2) = imgX(index2 - distance)
                imgY(index2) = imgY(index2 - distance)
                
                index2 = index2 - distance
                If index2 - distance < firstEl Then Exit Do
            Loop
            arr(index2) = value1
            imgItem(index2) = value2
            imgX(index2) = value3
            imgY(index2) = value4
        Next
    Loop Until distance = 1
End Sub

Private Sub Form_Resize()
    Me.Cls
    Call Init
End Sub
