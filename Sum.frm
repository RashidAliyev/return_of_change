VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Размелчение У.Е."
   ClientHeight    =   3375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5190
   Icon            =   "Sum.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   5190
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Напечатать"
      Height          =   345
      Left            =   1200
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdCalculation 
      Caption         =   "&Вычислить"
      Default         =   -1  'True
      Height          =   345
      Left            =   2520
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&О программе"
      Height          =   345
      Left            =   3840
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.VScrollBar Scroll 
      Height          =   345
      Left            =   720
      Max             =   99
      Min             =   1
      TabIndex        =   2
      Top             =   120
      Value           =   1
      Width           =   255
   End
   Begin VB.TextBox SumBox 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1049
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      MaxLength       =   2
      TabIndex        =   0
      Text            =   "1"
      Top             =   120
      Width           =   600
   End
   Begin VB.ListBox lstResult 
      Height          =   2400
      ItemData        =   "Sum.frx":0CCA
      Left            =   120
      List            =   "Sum.frx":0CCC
      TabIndex        =   3
      ToolTipText     =   "Кликните на поле чтоб скопировать содержимое строки в буфер обмена."
      Top             =   840
      Width           =   4935
   End
   Begin VB.CommandButton cmdURL 
      Caption         =   "&Web page"
      Height          =   345
      Left            =   1200
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1 уе/шт"
      Height          =   195
      Index           =   4
      Left            =   3000
      TabIndex        =   10
      Top             =   600
      Width           =   570
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "5 уе/шт"
      Height          =   195
      Index           =   3
      Left            =   2280
      TabIndex        =   9
      Top             =   600
      Width           =   570
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10 уе/шт"
      Height          =   195
      Index           =   2
      Left            =   1560
      TabIndex        =   8
      Top             =   600
      Width           =   660
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "20 уе/шт"
      Height          =   195
      Index           =   1
      Left            =   840
      TabIndex        =   7
      Top             =   600
      Width           =   660
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Версия"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   555
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public A, U, X, X1, Y, Y1, Z, Z1, K
Private Sub cmdAbout_Click()
    MsgBox "Sum " & App.Major & "." & App.Minor & vbCrLf & _
    "Автор: Рашид Алиев" & vbCrLf & vbCrLf & _
    "Прграмма создана на Visual Basic 6.0." & vbCrLf & vbCrLf & _
    "Сайт: https://RashidAliyev.Com" & vbCrLf & _
    "Э-почта: " & App.LegalCopyright, vbInformation, Mid$(cmdAbout.Caption, 2)
End Sub
Private Sub cmdCalculation_Click()
5 lstResult.Clear
6 K = 0
10 If A < 99 And A > 1 Then GoTo 30

20 If A > 99 Then MsgBox "Число не должно превышат 99!", vbExclamation, "Exclamation": Exit Sub
   If A < 1 Then MsgBox "Число должно быть больще 1!", vbExclamation, "Exclamation": Exit Sub

30 X1 = Int(A / 20)
   Y1 = Int(A / 10)
   Z1 = Int(A / 5)
For X = 0 To X1
For Y = 0 To Y1
For Z = 0 To Z1
For U = 0 To A
    F = 20 * X + 10 * Y + 5 * Z + U
    If F = A Then K = K + 1: lstResult.AddItem K & vbTab & X & vbTab & Y & vbTab & Z & vbTab & U
Next U, Z, Y, X
End Sub
Private Sub cmdPrint_Click()
Dim LI As Integer
'Напечатать расчёт
On Error GoTo ErrorHandler  'Set up error handler
If lstResult.ListCount < 1 Then
    MsgBox "Сначала выполните вычесление!", vbExclamation, "Внимание!"
    Exit Sub
End If
Printer.CurrentX = 10
Printer.FontName = "Courier New Cyr"
Printer.FontSize = "10"
Printer.Print " Rashid Aliyev - Sum " & App.Major & "." & App.Minor
Printer.Print "------------------------------------"
Printer.Print " Версия 20     10       5        1"
Printer.Print "------------------------------------"

For LI = 1 To lstResult.ListCount
    Printer.Print " " & lstResult.List(LI)
Next LI
    
Printer.EndDoc ' Print done

Exit Sub

ErrorHandler:
    MsgBox "Проплемы с печатью на ваш принтер!", vbCritical + vbMsgBoxHelpButton, "Ошибка"
    Exit Sub
End Sub
Private Sub cmdURL_Click()
    Shell "explorer http://www.rashid4ever.narod.ru/myapps/", vbMaximizedFocus
End Sub
Private Sub Form_Load()
K = 0
A = Val(SumBox.Text)
End Sub
Private Sub Form_Resize()
If Me.Height < 3780 Then Me.Height = 3780
'If Me.Height > 6000 Then Me.Height = 6000
If Me.Width < 5310 Then Me.Width = 5310
'If Me.Width > 10500 Then Me.Width = 10500
Me.lstResult.Height = Me.Height - 1300
Me.lstResult.Width = Me.Width - 380
Me.cmdAbout.Left = Me.Width - 2280 + 1215 - 380
Me.cmdCalculation.Left = Me.Width - 2280 + 1215 - 380 - 1335
Me.cmdPrint.Left = Me.Width - 2280 + 1215 - 380 - 1335 - 1335
If Me.Width >= 6660 Then cmdURL.Visible = True Else cmdURL.Visible = False
End Sub
Private Sub lstResult_Click()
Clipboard.Clear
Clipboard.SetText lstResult.List(lstResult.ListIndex)
End Sub
Private Sub Scroll_Change()
    SumBox.Text = Scroll.Value
    A = Val(SumBox.Text)
End Sub
Private Sub SumBox_Change()
On Error GoTo 10
Scroll.Value = Val(SumBox.Text): Exit Sub
10 SumBox.Text = Scroll.Value
End Sub
