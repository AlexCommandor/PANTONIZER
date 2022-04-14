VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PANTONIZER"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11175
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   11175
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Enter here desired LAB coordinates:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   3375
      Begin VB.TextBox txtLAB 
         Height          =   285
         Index           =   2
         Left            =   2160
         TabIndex        =   2
         Text            =   "0.00"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtLAB 
         Height          =   285
         Index           =   1
         Left            =   1200
         TabIndex        =   1
         Text            =   "0.00"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtLAB 
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   0
         Text            =   "0.00"
         Top             =   480
         Width           =   855
      End
   End
   Begin MSComctlLib.ListView lstPantones 
      Height          =   3855
      Index           =   0
      Left            =   360
      TabIndex        =   4
      Top             =   2160
      Visible         =   0   'False
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   6800
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "name"
         Text            =   "Name"
         Object.Width           =   5715
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Key             =   "labL"
         Text            =   "L"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Key             =   "labA"
         Text            =   "a"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Key             =   "labB"
         Text            =   "b"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "color"
         Text            =   "Color"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Key             =   "dE"
         Text            =   "dE"
         Object.Width           =   1905
      EndProperty
   End
   Begin MSComctlLib.TabStrip tabBooks 
      Height          =   4695
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   8281
      MultiRow        =   -1  'True
      Separators      =   -1  'True
      TabStyle        =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   1215
      Left            =   3960
      Top             =   120
      Width           =   6975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   120
      Top             =   6360
      Width           =   10935
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const LOCALE_SDECIMAL = &HE         '  decimal separator

Private Declare Function GetUserDefaultLCID% Lib "kernel32" ()
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" _
  (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long


Dim sTables() As Variant, iNumTables As Integer, vRGB As tRGB, vLAB As tCIELab, vLABnew As tCIELab ', sSeparator As String

Private Sub Form_Load()
    Dim i As Integer, j As Integer, arrData As Variant, arrTable As Variant
    Dim sPantone() As String
    
    On Error Resume Next
    
    iNumTables = 0
    
    Me.tabBooks.Tabs.Clear
    
    For i = 101 To 199
        arrData = LoadResData(i, "CUSTOM")
        If Err.Number = 0 Then
            arrTable = StrConv(arrData, vbUnicode)
            arrTable = Replace$(arrTable, ".", GetDecimalSeparator)
            iNumTables = iNumTables + 1
            ReDim Preserve sTables(1 To iNumTables)
            sTables(iNumTables) = Split(arrTable, vbCrLf)
            Me.tabBooks.Tabs.Add , , sTables(iNumTables)(0)
        Else
            Err.Clear
            Exit For
        End If
    Next i
    
    On Error GoTo 0
    
'    Load frmSplash
'    frmSplash.Show
    
    Dim liRow As ListItem, vSorting(1 To 3) As Variant
    For i = 1 To iNumTables
        Load Me.lstPantones(i)
        For j = 1 To UBound(sTables(i))
            sPantone = Split(sTables(i)(j), vbTab)
            Set liRow = Me.lstPantones(i).ListItems.Add(sPantone(0), sPantone(1), sPantone(1))
            liRow.ListSubItems.Add 1, "labL", sPantone(2)
            liRow.ListSubItems.Add 2, "labA", sPantone(3)
            liRow.ListSubItems.Add 3, "labB", sPantone(4)
            liRow.ListSubItems.Add 4, "colorRGB", String(15, Chr$(&H7F))
            Call CieLAB_RGB(sPantone(2), sPantone(3), sPantone(4), vRGB.R, vRGB.G, vRGB.B)
            liRow.ListSubItems(4).ForeColor = RGB(vRGB.R, vRGB.G, vRGB.B)
            liRow.ListSubItems(4).Bold = True
            liRow.ListSubItems.Add 5, "dE", vbNullString
        Next j
        Me.lstPantones(i).Visible = True
    Next i
    Me.lstPantones(1).ZOrder 0
    'UpdateDeltaE
'    frmSplash.Hide
'    Unload frmSplash
End Sub

Private Sub lstPantones_ColumnClick(Index As Integer, ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Me.lstPantones(Index).SortKey = ColumnHeader.SubItemIndex
    'Me.lstPantones(Index).SortKey = 0
    Me.lstPantones(Index).SortOrder = 1 - Me.lstPantones(Index).SortOrder
    Me.lstPantones(Index).Sorted = True
End Sub

Private Sub lstPantones_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
    Me.Shape1.FillColor = Item.ListSubItems(4).ForeColor
    Me.Shape1.Refresh
End Sub

Private Sub tabBooks_Click()
    Me.lstPantones(Me.tabBooks.SelectedItem.Index).ZOrder 0
End Sub

Private Function SortAsc(ByRef vArray As Variant, Optional ByVal iArrayIndex As Integer = 1) As Variant
    Dim iRes() As Integer, i As Integer, j As Integer, iU As Integer, iTmp As Integer
    Dim sTmp1() As String, sTmp2() As String
    If Not IsArray(vArray) Then Exit Function
    iU = UBound(vArray)
    ReDim iRes(1 To iU)
    For i = 1 To iU
        iRes(i) = i
    Next i
    
    For j = 1 To iU
        For i = 2 To iU
            sTmp1 = Split(vArray(iRes(i - 1)), vbTab)
            sTmp2 = Split(vArray(iRes(i)), vbTab)
            If Val(sTmp2(iArrayIndex)) < Val(sTmp1(iArrayIndex)) Then
                iTmp = iRes(i)
                iRes(i) = iRes(i - 1)
                iRes(i - 1) = iTmp
            End If
        Next i
        DoEvents
    Next j
    SortAsc = iRes
End Function

Private Sub txtLAB_Change(Index As Integer)
    If Val(Me.txtLAB(Index).Text) > 100 Then Me.txtLAB(Index).Text = "100.00"
    If Index = 0 Then
        If Val(Me.txtLAB(Index).Text) < 0 Then Me.txtLAB(Index).Text = "0.00"
    Else
        If Val(Me.txtLAB(Index).Text) < -100 Then Me.txtLAB(Index).Text = "-100.00"
    End If
    Call CieLAB_RGB(Val(txtLAB(0).Text), Val(txtLAB(1).Text), Val(txtLAB(2).Text), vRGB.R, vRGB.G, vRGB.B)
    Me.Shape2.FillColor = RGB(vRGB.R, vRGB.G, vRGB.B)
End Sub

Private Sub txtLAB_GotFocus(Index As Integer)
    Me.txtLAB(Index).SelStart = 0
    Me.txtLAB(Index).SelLength = Len(Me.txtLAB(Index).Text)
    vLAB.L = Val(Me.txtLAB(0).Text)
    vLAB.A = Val(Me.txtLAB(1).Text)
    vLAB.B = Val(Me.txtLAB(2).Text)
End Sub

Private Sub txtLAB_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 44 Then KeyAscii = 46
    If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 46) And (KeyAscii <> 45) Then KeyAscii = 0
    If (KeyAscii = 45) And (Index = 0) Then KeyAscii = 0
End Sub

Private Sub txtLAB_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38 'UP
            Me.txtLAB(Index).Text = Format$(Val(txtLAB(Index).Text) + 1 + 9 * Shift, "##0.00")
        Case 40 'DOWN
            Me.txtLAB(Index).Text = Format$(Val(txtLAB(Index).Text) - 1 - 9 * Shift, "##0.00")
    End Select
End Sub

Private Sub txtLAB_LostFocus(Index As Integer)
    txtLAB(Index).Text = Format$(Val(txtLAB(Index).Text), "##0.00")
    vLABnew.L = Val(Me.txtLAB(0).Text)
    vLABnew.A = Val(Me.txtLAB(1).Text)
    vLABnew.B = Val(Me.txtLAB(2).Text)
    If (vLAB.L <> vLABnew.L) Or (vLAB.A <> vLABnew.A) Or (vLAB.B <> vLABnew.B) Then
        vLAB = vLABnew
        Call UpdateDeltaE
    End If
End Sub

Private Sub UpdateDeltaE()
    Dim i As Integer, j As Integer, deltaE As Double, sDelta As String
    For i = 1 To iNumTables
        For j = 1 To Me.lstPantones(i).ListItems.Count
            deltaE = Sqr( _
                (vLAB.L - Val(Me.lstPantones(i).ListItems(j).ListSubItems(1).Text)) ^ 2 + _
                (vLAB.A - Val(Me.lstPantones(i).ListItems(j).ListSubItems(2).Text)) ^ 2 + _
                (vLAB.B - Val(Me.lstPantones(i).ListItems(j).ListSubItems(3).Text)) ^ 2 _
                        )
            sDelta = Format$(deltaE, "###0.00")
            sDelta = Space$(10 - Len(sDelta)) & sDelta
            Me.lstPantones(i).ListItems(j).ListSubItems(5).Text = sDelta
            If Val(sDelta) < 3 Then
                Me.lstPantones(i).ListItems(j).ListSubItems(5).Bold = True
                Me.lstPantones(i).ListItems(j).ListSubItems(5).ForeColor = RGB(0, 128, 0)
            ElseIf Val(sDelta) < 6 Then
                Me.lstPantones(i).ListItems(j).ListSubItems(5).Bold = True
                Me.lstPantones(i).ListItems(j).ListSubItems(5).ForeColor = vbBlue
            Else
                Me.lstPantones(i).ListItems(j).ListSubItems(5).Bold = False
                Me.lstPantones(i).ListItems(j).ListSubItems(5).ForeColor = vbRed
            End If
        Next j
    Me.lstPantones(i).SortKey = 5
    Me.lstPantones(i).SortOrder = lvwAscending
    Me.lstPantones(i).Sorted = True
    Next i
End Sub

Private Function GetDecimalSeparator() As String
  Dim iLocale As Integer, sTmpStr As String, lRes As Long, aLen As Long
  On Error Resume Next
  sTmpStr = String$(255, " ") & Chr$(0)
  aLen = 1
  iLocale = GetUserDefaultLCID()
  lRes = GetLocaleInfo(iLocale, LOCALE_SDECIMAL, sTmpStr, aLen)
  GetDecimalSeparator = Left$(sTmpStr, aLen)
  Err.Clear
  On Error GoTo 0
End Function

