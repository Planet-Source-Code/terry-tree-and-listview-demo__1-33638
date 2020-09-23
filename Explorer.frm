VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ExplorerForm 
   Caption         =   "TreeView and ListView Control Demo"
   ClientHeight    =   5085
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   ScaleHeight     =   5085
   ScaleWidth      =   8880
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView ListView1 
      Height          =   4440
      Left            =   3600
      TabIndex        =   2
      Top             =   150
      Width           =   5190
      _ExtentX        =   9155
      _ExtentY        =   7832
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   4440
      Left            =   120
      TabIndex        =   1
      Top             =   150
      Width           =   3360
      _ExtentX        =   5927
      _ExtentY        =   7832
      _Version        =   393217
      Style           =   7
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   -15
      Top             =   4680
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3615
      TabIndex        =   0
      Top             =   4665
      Width           =   5175
   End
End
Attribute VB_Name = "ExplorerForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim FSys As FileSystemObject
Dim TimerBusy As Boolean

Sub ScanFolder(folderSpec)
Dim thisFolder As Folder
Dim allFolders As Folders
    
    Set thisFolder = FSys.GetFolder(folderSpec)
    Set allFolders = thisFolder.SubFolders
    For Each thisFolder In allFolders
        TreeView1.Nodes.Add UCase(thisFolder.ParentFolder.Path), tvwChild, UCase(thisFolder.Path), thisFolder.Name
        ScanFolder (thisFolder.Path)
    Next
End Sub

Private Sub Form_Load()
    LWidth = ListView1.Width - 5 * Screen.TwipsPerPixelX
    ListView1.ColumnHeaders.Add 1, , "File Name", 0.3 * LWidth
    ListView1.ColumnHeaders.Add 2, , "Size", 0.2 * LWidth, lvwColumnRight
    ListView1.ColumnHeaders.Add 3, , "Created", 0.25 * LWidth
    ListView1.ColumnHeaders.Add 4, , "Modified", 0.25 * LWidth
    Set FSys = CreateObject("Scripting.FileSystemObject")
' Specify the folder you wish to map in the following line
    InitPath = "C:\WINDOWS"
    TreeView1.Nodes.Add , tvwFirst, UCase(InitPath), InitPath
    Me.Show
    Screen.MousePointer = vbHourglass
    DoEvents
    ScanFolder (InitPath)
    Screen.MousePointer = vbDefault
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListView1.SortKey = ColumnHeader.Index - 1
    ' Set Sorted to True to sort the list.
    ListView1.Sorted = True
End Sub

Private Sub Timer1_Timer()
'   If this event handler is executing, don't start it again
    If TimerBusy Then Exit Sub
'   Timerbusy indicates that the subroutine is calculating
    TimerBusy = True
    Dim totFiles As Integer, totSize As Long
    totFiles = 0
    totSize = 0
    On Error Resume Next
'   scan all files in the ListView control and examine
'   their Selected property. If it's True, then increase the
'   count of selected files (totFiles) and add its size to the total (TotSize)
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Selected Then
            totFiles = totFiles + 1
            totSize = totSize + ListView1.ListItems(i).SubItems(1)
        End If
    Next
'   Display the file count and their total size
    Label1.Caption = Format(totFiles, "###,##0") & " files selected, containing " & Format(totSize, "###,###,###,##0") & " bytes"
'   and reset the TimerBusy variable, so that it can be invoked again
    TimerBusy = False
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
' When a folder is selected on the TreeView control, its files
' must be displayed in the ListView control.
' To display the files and their attributes, it scans the
' members of the allFiles collection, which contains all the files
' in the selected folder
Dim thisFolder As Folder
Dim thisFile As File
Dim allFiles As Files
Dim thisItem  As ListItem

    Screen.MousePointer = vbHourglass
    ListView1.ListItems.Clear
'   Create a Folder variable that references the selected folder
    Set thisFolder = FSys.GetFolder(Node.Key)
'   and use its Files property to retrieve the folder's files
    Set allFiles = thisFolder.Files
    If allFiles.Count > 0 Then
On Error Resume Next
'   Now scan all the files in the allFiles collection
'   Use the properties of the variable thisItem to retrieve the file attributes
'   and attach them to the ListView control as subitems of the current file
        For Each thisFile In allFiles
                Set thisItem = ListView1.ListItems.Add(, , thisFile.Name)
                thisItem.SubItems(1) = Format(thisFile.Size, "###,###,###")
                thisItem.SubItems(2) = Left(thisFile.DateCreated, 8)
                thisItem.SubItems(3) = Left(thisFile.DateLastModified, 8)
                If thisFile.Attributes And vbSystem Then thisItem.Ghosted = True
        Next
    End If
    Screen.MousePointer = vbDefault

End Sub
