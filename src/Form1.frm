VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DirMaker"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   6765
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView ListView1 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   7646
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.CommandButton Command3 
      Caption         =   "+"
      Height          =   375
      Left            =   6240
      TabIndex        =   1
      Top             =   4920
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public stdIn As String

Private Function PreselectItems() As Boolean
    UnselectAll
    Dim Item As ListItem
    For Each Item In ListView1.ListItems
        If Item = "dwg" Or Item = "map" Or Item = "src" Or Item = "pdf" Then
            Item.Selected = True
            End If
    Next Item
End Function

Private Function UnselectAll() As Boolean
    Dim Item As ListItem
    For Each Item In ListView1.ListItems
            Item.Selected = False
    Next Item
End Function

Private Function FolderExists(sFullPath As String) As Boolean
    Dim myFSO As Object
    Set myFSO = CreateObject("Scripting.FileSystemObject")
    FolderExists = myFSO.FolderExists(sFullPath)
End Function

Private Function AddDir(dirName As String) As Boolean
    Dim fullPath As String
    fullPath = stdIn + "\" + dirName
    Debug.Print fullPath
    If FolderExists(fullPath) = False Then
        MkDir fullPath
    End If
End Function

Private Function UpdateList() As Boolean
    Dim Item As ListItem
    For Each Item In ListView1.ListItems
        If FolderExists(stdIn + "\" + Item.Text) = True Then
            Item.Checked = True
        Else
            Item.Checked = False
        End If
    Next Item
End Function

Private Sub Command3_Click()
    Dim Item As ListItem
    For Each Item In ListView1.ListItems
        If Item.Selected Then
            AddDir (Item)
        End If
    Next Item
    UpdateList
    UnselectAll
End Sub

'Store command line arguments in this array
   ' Dim sArgs() As String

    
  '  Dim iLoop As Integer
    'Assuming that the arguments passed from
    'command line will have space in between,
    'you can also use comma or otehr things...
   ' sArgs = Split(Command$, " ")
   ' For iLoop = 0 To UBound(sArgs)
        'this will print the command line
        'arguments that are passed from the command line
        'Debug.Print sArgs(iLoop)
    'Next

Private Sub Form_Load()
    stdIn = Replace(Command, Chr(34), "")
    Label1.Caption = stdIn
    If FolderExists(stdIn) = False Then
        Label1.ForeColor = vbRed
        Label1.Caption = "Ошибка пути. Нет такой папки: " + Label1.Caption
        ListView1.Enabled = False
        Command3.Enabled = False
    End If
    Debug.Print stdIn

    ' Dim MyDict As New Dictionary(Of String, List(Of Double))
    ' MyDict.Add("One", New List(Of Double)(New Double() {1.3, 2.4, 6.9}))
    ' MyDict.Add("Two", New List(Of Double)(New Double() {2.4, 45.4, 9}))
    ' MyDict.Add("Three", New List(Of Double)(New Double() {3.5, 2.4, 16.9}))
    Dim Dict As Dictionary
    Set Dict = New Dictionary
    Dict.Add "axis", "Акты разбивки"
    Dict.Add "credo", "Файлы CredoDAT"
    Dict.Add "doc", "Документы по объекту"
    Dict.Add "dwg", "Autocad DWG"
    Dict.Add "egrn", "Выписки ЕГРН"
    Dict.Add "foto", "Фотографии"
    Dict.Add "kpt", "Выписки КПТ"
    Dict.Add "map", "MapInfo"
    Dict.Add "pdf", "PDF файлы для печати"
    Dict.Add "proj", "Проектная документация"
    Dict.Add "sat", "Спутниковые снимки"
    Dict.Add "src", "Полевые измерения"

    Debug.Print "List clicked"
    
    With ListView1
        Dim itmX As ListItem ' Create a variable to add ListItem objects.
        Dim clmX As ColumnHeader ' Create an object variable for the ColumnHeader object.
        ' Add ColumnHeaders.
        Set clmX = .ColumnHeaders.Add(, , "Папка", .Width / 5)
        Set clmX = .ColumnHeaders.Add(, , "Описание", .Width * 4 / 5)
        .BorderStyle = ccFixedSingle ' Set BorderStyle property.
        .View = lvwReport ' Set View property to Report.
    
        Dim key As Variant
        For Each key In Dict.Keys
            Debug.Print key + "  " + Dict.Item(key)
            ' Add a main item
            Set itmX = .ListItems.Add(, , key)
            ' Add two subitems for that item
            itmX.SubItems(1) = Dict.Item(key)
            'itmX.ForeColor = vbGreen
        Next key
   End With
   UpdateList
   PreselectItems
End Sub

'Private Sub ListView1_ItemDblClick(ByVal Item As MSComctlLib.ListItem)
'Label2.Caption = ListView1.SelectedItem.Text
'End Sub

Private Sub ListView1_DblClick()
    'Label2.Caption = ListView1.SelectedItem.Text
    AddDir (ListView1.SelectedItem.Text)
    UpdateList
    UnselectAll
End Sub

Private Sub Form_Paint()
    UpdateList
End Sub

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Item.Checked = Not Item.Checked
End Sub
