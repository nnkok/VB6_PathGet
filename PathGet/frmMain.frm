VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13215
   LinkTopic       =   "Form1"
   ScaleHeight     =   6810
   ScaleWidth      =   13215
   StartUpPosition =   3  'Windows の既定値
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   315
      Left            =   1020
      TabIndex        =   6
      Top             =   2700
      Width           =   735
   End
   Begin MSComDlg.CommonDialog ComDialog 
      Left            =   120
      Top             =   5460
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Start_Cmd 
      Caption         =   "指定パス　以下のファイルを全て表示"
      Height          =   735
      Left            =   60
      TabIndex        =   4
      Top             =   6060
      Width           =   13035
   End
   Begin VB.ListBox File_List 
      Height          =   2940
      Left            =   60
      TabIndex        =   3
      Top             =   3060
      Width           =   13035
   End
   Begin VB.TextBox Folder_Path 
      Height          =   315
      Left            =   60
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2220
      Width           =   13035
   End
   Begin VB.DriveListBox Drive_List 
      Height          =   300
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   13035
   End
   Begin VB.DirListBox Dir_List 
      Height          =   1770
      Left            =   60
      TabIndex        =   0
      Top             =   420
      Width           =   13035
   End
   Begin VB.Label Lbl_Count 
      Caption         =   "Label1"
      Height          =   315
      Left            =   60
      TabIndex        =   5
      Top             =   2700
      Width           =   3915
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    With ComDialog
        .CancelError = True
        .DialogTitle = "保存されている範囲名から選択して下さい"
        .InitDir = "C\"
        .Filter = "範囲ファイル(Area)|*.AREA|all files|*.*"
        .Flags = cdlOFNFileMustExist Or cdlOFNLongNames Or cdlOFNExplorer Or cdlOFNNoChangeDir
        .ShowOpen       'ここでカレントフォルダがApp.Pathから変わってしまうとgb32がこれ以降見つからなくなってしまうからcdlOFNNoChangeDirは必須
'        sInp = .FileName
    End With
End Sub

Private Sub Dir_List_Click()
    Folder_Path.Text = Dir_List.Path
End Sub

Private Sub Drive_List_Change()
    Dir_List.Path = Drive_List
End Sub

Private Sub Form_Resize()
    
    Dim dblWide    As Double
    
    dblWide = frmMain.Width - 250

    Drive_List.Width = dblWide
    Dir_List.Width = dblWide
    Folder_Path.Width = dblWide
    File_List.Width = dblWide
    Start_Cmd.Width = dblWide

    If frmMain.Height < 7200 Then frmMain.Height = 7200
    If frmMain.Height > 7200 Then frmMain.Height = 7200

End Sub

Private Sub Start_Cmd_Click()
    Dim i As Long
    Call GetStart(Folder_Path.Text)
    If 0 < UBound(stfilePath) Then
        For i = 0 To UBound(stfilePath) Step 1
            File_List.AddItem stfilePath(i)
            DoEvents
        Next
    End If
End Sub
