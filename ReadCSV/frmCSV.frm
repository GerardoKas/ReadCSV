VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   7065
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frBottom 
      BorderStyle     =   0  'None
      Height          =   510
      Left            =   0
      TabIndex        =   7
      Top             =   2610
      Width           =   6360
      Begin VB.TextBox Text2 
         Height          =   330
         Left            =   0
         TabIndex        =   11
         Text            =   "Buscar"
         Top             =   90
         Width           =   2085
      End
      Begin VB.CommandButton findText 
         Caption         =   "Marcar"
         Height          =   330
         Left            =   2115
         TabIndex        =   10
         Top             =   90
         Width           =   780
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Notepad"
         Height          =   285
         Left            =   4275
         TabIndex        =   9
         Top             =   135
         Width           =   1005
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "STOP"
         Height          =   285
         Left            =   5400
         TabIndex        =   8
         Top             =   135
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Buscados"
         Height          =   195
         Left            =   2925
         TabIndex        =   12
         Top             =   135
         Width           =   705
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "TABS"
      Height          =   285
      Left            =   4455
      TabIndex        =   6
      Top             =   45
      Width           =   690
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmCSV.frx":0000
      Left            =   5265
      List            =   "frmCSV.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   45
      Width           =   1680
   End
   Begin VB.CommandButton Command3 
      Caption         =   "EQUALS"
      Height          =   300
      Left            =   3600
      TabIndex        =   4
      Top             =   45
      Width           =   825
   End
   Begin VB.CommandButton Command2 
      Caption         =   "LINES"
      Height          =   300
      Left            =   2790
      TabIndex        =   3
      Top             =   50
      Width           =   780
   End
   Begin MSComctlLib.ListView lv1 
      Height          =   2175
      Left            =   0
      TabIndex        =   2
      Top             =   450
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   3836
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CSV"
      Height          =   300
      Left            =   2070
      TabIndex        =   1
      Top             =   50
      Width           =   690
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   0
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   50
      Width           =   1995
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim items As Collection
Dim itsLine As Collection
Dim fileData As String
Dim equalStr As String
Dim rx As RegExp
Dim bSalir As Boolean

Const titulo As String = "Proyecto CSV - GCM 2006"

Private Declare Function OemToChar Lib "user32" Alias "OemToCharA" (ByVal lpszSrc As String, ByVal lpszDst As String) As Long
Private Declare Function CharToOem Lib "user32" Alias "CharToOemA" (ByVal lpszSrc As String, ByVal lpszDst As String) As Long
'Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'Private Const SW_SHOWDEFAULT = 10
'Private Const SW_SHOWMAXIMIZED = 3
'Private Const SW_SHOWMINIMIZED = 2
'Private Const SW_SHOWMINNOACTIVE = 7
'Private Const SW_SHOWNOACTIVATE = 4
'Private Const SW_SHOWNORMAL = 1

Private Const iniFile = "Comandos.ini"

Private Sub cmdStop_Click()
bSalir = True
End Sub

Private Sub Combo1_Click()
execCommand Combo1.Text
End Sub

'CSV
Private Sub Command1_Click()
Dim ff As TextStream, sys As FileSystemObject
Set sys = New FileSystemObject

If checkFile = False Then doNoFile: Exit Sub

Set ff = sys.OpenTextFile(Text1.Text)
lv1.ListItems.Clear
lv1.ColumnHeaders.Clear
On Local Error GoTo notis
Do
cabecera = Dos2Win(ff.ReadLine)
Loop While cabecera = ""

getCSVItems (cabecera)
addColumns items
Do
    If doStop Then Exit Sub
    linea = Dos2Win(ff.ReadLine)
    getCSVItems (linea)
    addline items
Loop While Not ff.AtEndOfStream


Exit Sub
notis:
MsgBox "No funciona con esteformato (CSV)" & vbCrLf & showError(Err)
End Sub

'LINES
Private Sub Command2_Click()
Dim line As String
Dim ff As TextStream
Dim sys As FileSystemObject
Set sys = New FileSystemObject

If checkFile = False Then doNoFile: Exit Sub

Set ff = sys.OpenTextFile(Text1.Text)
lv1.ListItems.Clear
lv1.ColumnHeaders.Clear
Set itsLine = New Collection
'On Local Error GoTo notis
Do
line = (ff.ReadLine)
If doStop Then Exit Sub
If isDivisor(line) Then
    'nuevo servicio
    If itsLine.count > 0 Then
    addLineItem itsLine
    Set itsLine = New Collection
    End If
Else
itsLine.Add line
End If
Me.Caption = "Linea: " & ff.line
Loop While Not ff.AtEndOfStream
ff.Close


Exit Sub
notis:
MsgBox "No funciona con esteformato (LINES)" & vbCrLf & showError(Err)
End Sub

Sub addLineItem(col As Collection)
Dim pts, eqSep
Dim noHeaders As Boolean
Dim li As ListItem
If lv1.ColumnHeaders.count = 0 Then noHeaders = True
For i = 1 To col.count
    eqSep = getEqualSeparator(col(i))
    If eqSep = "" Then GoTo DoNext
    pts = Split(col(i), eqSep)
    If UBound(pts) < 1 Then
        'li.SubItems(i - 2) = li.SubItems(i - 2) & Trim(pts(0))
        ReDim Preserve pts(1)
        pts(1) = Trim(pts(0))
        pts(0) = "ANY"
    Else
    pts(0) = Trim(pts(0)) 'titulo
    pts(1) = Trim(pts(1)) 'valor
    End If
    On Local Error Resume Next
    columna = lv1.ColumnHeaders.Item(pts(0)).Index
    If Err.Number = 35601 Then
        'MsgBox Err.Description
        lv1.ColumnHeaders.Add , pts(0), pts(0)
        columna = lv1.ColumnHeaders.Item(pts(0)).Index
        'On Error GoTo otError
        On Error GoTo 0
    End If
    If i = 1 Then 'crear listitem
        Set li = lv1.ListItems.Add(, , pts(1))
    Else
        li.SubItems(columna - 1) = pts(1)
    End If
DoNext:
Next
Exit Sub
otError:
MsgBox showError(Err)
End Sub

'EQUAL
Private Sub Command3_Click()
Dim ff As TextStream
Dim sys As FileSystemObject
Set sys = New FileSystemObject

If checkFile = False Then doNoFile: Exit Sub

Set ff = sys.OpenTextFile(Text1.Text)
lv1.ListItems.Clear
lv1.ColumnHeaders.Clear
On Local Error GoTo notis
Dim cero As Boolean
cero = True

Do
    If doStop Then Exit Sub
    linea = (ff.ReadLine)
    If linea <> "" Then
        getEqualSeparator (linea)
        If cero = True Then
            addEqualColumns (2)
            cero = False
        End If
        addEqualItems (linea)
    End If
Loop While Not ff.AtEndOfStream


Exit Sub
notis:
MsgBox "No funciona con esteformato (EQUAL)" & vbCrLf & showError(Err)
End Sub

Function getEqualSeparator(line As String) As String
cad = Array("=", ":", ",", "|", "\")
getEqualSeparator = ""
equalStr = ""
For i = 0 To UBound(cad)
    p = Split(line, cad(i))
    If UBound(p) >= 1 Then
        done = True
        getEqualSeparator = cad(i)
        equalStr = cad(i)
        Exit Function
    End If
Next
End Function

Function isDivisor(line As String) As Boolean
Set rx = New RegExp
'si es una linea toda igual
rx.Pattern = "^(\W+)$"
If rx.Test(line) = True Or line = "" Then isDivisor = True
End Function

Sub addEqualItems(line)
Dim p, li As ListItem
'On Error Resume Next
p = Split(line, equalStr, 2)
Set li = lv1.ListItems.Add(, , p(0))
li.SubItems(1) = p(1)
End Sub

Sub addEqualColumns(Number)
For i = 1 To Number
    lv1.ColumnHeaders.Add i, , "Col" & i
Next
End Sub

'TABS
Private Sub Command4_Click()
Dim deletePrevio As Boolean
Dim ff As TextStream
Dim sys As FileSystemObject
Set sys = New FileSystemObject

If checkFile = False Then doNoFile: Exit Sub

Set ff = sys.OpenTextFile(Text1.Text)
lv1.ListItems.Clear
lv1.ColumnHeaders.Clear
'On Error GoTo notis
Set items = New Collection
Dim numCols As Integer
Do
    cabecera = (ff.ReadLine)
    If cabecera <> "" Then getTABItems (cabecera)
Loop While items.count <= 1

numCols = items.count
addColumns items
Do
    If doStop Then Exit Sub
    linea = (ff.ReadLine)
    If linea <> "" Then
    If iniciaTAB(linea) = False Then
        Set items = New Collection
        deletePrevio = False
    Else
        deletePrevio = True
    End If
    getTABItems (linea)
    If items.count <= numCols And items.count > 1 Then
    addline items, deletePrevio
    End If
    End If
Loop While Not ff.AtEndOfStream


Exit Sub
notis:
MsgBox "No funciona con esteformato" & vbCrLf & showError(Err)
End Sub

Function iniciaTAB(line)
If InStr(1, line, "  ") = 1 Or InStr(1, line, vbTab) = 1 Then
    iniciaTAB = True
Else
    iniciaTAB = False
End If
End Function
Sub getTABItems(line As String)
Static oldPieces
Static numPieces
Dim mts As MatchCollection
Dim mt As Match
Set rx = New RegExp
'si es una linea toda igual
If isDivisor(line) Then Exit Sub

'buscar piezas
rx.Pattern = "(\x09|  )+": rx.Global = True
line2 = rx.Replace(line, "|*|")
myits = Split(line2, "|*|")
For i = 0 To UBound(myits)
    items.Add Trim(myits(i))
Next
If numPieces = 0 Then numPieces = UBound(myits) + 1
'si no hubo piezas .. y si no era tabulado
'If rx.Test(line) = False Then
    'algo no funciona, se supone que es texto
    'vamos a ver si contiene algo
'''    Set mt = mts.Item(1)
'''    MsgBox mt.FirstIndex
'''End If
If mts Is Nothing Then
    Set mts = rx.Execute(line)
    Debug.Print "POSES <" & line & ">"
    For Each mt In mts
    Debug.Print "POS:" & mt.FirstIndex + mt.Length
    Next
End If
End Sub

Private Sub Command5_Click()
Shell "notepad.exe " & fileData, vbNormalFocus
End Sub

Private Sub Form_Load()
Text1.Text = ""

loadCommands
End Sub

Sub addColumns(its As Collection)
For i = 1 To its.count
lv1.ColumnHeaders.Add , , its(i)
Next
End Sub

Sub addline(its As Collection, Optional deletePrevio As Boolean = False)
Dim li As ListItem
If deletePrevio = True Then
    lv1.ListItems.Remove (lv1.ListItems.count)
End If
Set li = lv1.ListItems.Add(, , its(1))
For i = 2 To its.count
    li.SubItems(i - 1) = its(i)
Next
End Sub
Function getCSVItems(line As String)
Dim rx As RegExp
Dim mts As MatchCollection
Dim mt As Match
Set items = New Collection
Set rx = New RegExp
rx.Pattern = """([^""]+)"""
rx.Global = True
Set mts = rx.Execute(line)
getCSVItems = mts.count
For i = 0 To getCSVItems - 1
'Set mt = mts.Item(i)
    t = mts.Item(i).Value
    t = Replace(t, """", "")
    items.Add t
Next
End Function

Function Dos2Win(line) As String
Dos2Win = String(Len(line), "#")
OemToChar line, Dos2Win
End Function


Private Sub Form_Resize()
If Me.WindowState = vbMinimized Then Exit Sub
Me.ScaleMode = vbPixels
esp = 3
'lv1.Left = esp
lv1.Top = Text1.Height + Text1.Top + esp
lv1.Width = Me.ScaleWidth - lv1.Left - esp
lv1.Height = Me.ScaleHeight - Text1.Height - Text1.Top - esp - (frBottom.Height)
frBottom.Top = Me.ScaleHeight - frBottom.Height
End Sub

Private Sub lv1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
lv1.SortKey = ColumnHeader.Index - 1
lv1.SortOrder = (1 Xor lv1.SortOrder)
lv1.SortKey
End Sub

Private Sub Text1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Text1.Text = Data.Files(1)
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    findText_Click
    findText.SetFocus
End If
    
End Sub

Sub loadCommands()
Dim sects
Combo1.Clear
ReadWrite.setIniFile ".\" & iniFile
ss = ReadWrite.ReadSections
sects = Split(ss, Chr(0))
For i = 0 To UBound(sects)
If sects(i) <> "" And sects(i) <> "MAIN" Then
    Combo1.AddItem sects(i)
End If
Next
End Sub

Sub execCommand(seccion As String)
    Dim errFile As String
    Dim wshshell As Object
    
    Set wshshell = CreateObject("WScript.Shell")
    
    cmd = ReadWrite.ReadFromFile(seccion, "cmd")
    
    output = ReadWrite.ReadFromFile("MAIN", "output")
    If output = "" Then output = "%CD%\TEMP.TXT"
    output = wshshell.ExpandEnvironmentStrings(output)
    errFile = output & ".ERR"
    
    If Dir$(output, vbNormal) <> "" Then
    Kill (output)
    End If
    If Dir$(errFile, vbNormal) <> "" Then
    Kill (errFile)
    End If

    comando$ = "cmd.exe /C " & cmd & " 1>" & output & " 2>" & errFile
    'ShellExec comando
    Me.Enabled = False
    r = wshshell.Run(comando, 5, True)
    Me.Enabled = True
    If FileLen(errFile) > 0 Then
        Shell "notepad.exe " & errFile, vbNormalFocus
    End If
    If FileLen(output) = 0 Then
        MsgBox "No se ha producido Salida Para el siguiente comando : " & vbCrLf & comando
        Exit Sub
    End If
    Text1.Text = output
    fileData = output
    tipo = ReadWrite.ReadFromFile(seccion, "TYPE")
    Select Case tipo
    Case "CSV"
        Command1_Click
    Case "LINE"
        Command2_Click
    Case "EQUAL"
        Command3_Click
    Case "TABS"
        Command4_Click
    End Select
    
End Sub
'FIND ITEMS
Private Sub findText_Click()
Dim line As String
Dim count As Long
For i = 1 To lv1.ListItems.count
    lv1.ListItems(i).Bold = False
    lv1.ListItems(i).ForeColor = vbBlack
    line = lv1.ListItems(i)
    For j = 1 To lv1.ColumnHeaders.count - 1
        line = line & vbTab & lv1.ListItems(i).SubItems(j)
    Next
    If InStr(1, line, Text2.Text, vbTextCompare) > 0 Then
        lv1.ListItems(i).ForeColor = vbRed
        lv1.ListItems(i).Bold = True
        count = count + 1
    End If
Next
Label1.Caption = count & " Encontrados [" & Text2.Text & "]"
lv1.SetFocus
End Sub


Function showError(tErr As ErrObject) As String
Dim t As String
showError = tErr.Number & vbCrLf & tErr.Description & vbCrLf & tErr.Source


End Function

Function doStop() As Boolean
DoEvents
If bSalir = True Then
    doStop = True
End If
bSalir = False
End Function

Function checkFile()
If Text1.Text = "" Then
    checkFile = False
Else
    checkFile = True
End If
End Function

Sub doNoFile()
MsgBox "El fichero " & vbrlf & Text1.Text & vbCrLf & "No Existe"
End Sub
