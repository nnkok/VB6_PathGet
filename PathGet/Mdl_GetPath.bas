Attribute VB_Name = "Mdl_GetPath"
Option Explicit

Public stfilePath() As String   '取得したファイルパスを全て配列化

'******************************************************************
'関数名：GetStart
'引数　：stFolderPath
'戻り値：なし
'機能　：指定されたフォルダーに格納されている全てのファイルのパスを配列化
'******************************************************************
Public Function GetStart(ByVal stFolderPath As String) As Boolean

    Dim lIndex          As Long
    Dim nPrompt         As String
    Dim nFilePath       As String
    Dim nFilePathes()   As String
'    Dim TextPath        As String
'    Dim FileNum         As Variant
    Dim lCont           As Long

    '現在"*.*"を指定することで全てのファイルが対象(変更することで拡張子の指定など可能　例）"*.txt"　でTEXTファイルの検索)
    nFilePathes() = GetFilesMostDeep(stFolderPath, "*.*")
'    TextPath = App.Path
'    FileNum = FreeFile
'    Open TextPath & "\Test.txt" For Output As #FileNum

    lCont = 0
    For lIndex = 1 To UBound(nFilePathes())
        ReDim Preserve stfilePath(lCont)
        stfilePath(lCont) = nFilePathes(lIndex)
'        nPrompt = nPrompt & nFilePathes(lIndex) & vbNewLine
        DoEvents
        lCont = lCont + 1
    Next lIndex
    
'    Print #FileNum, nPrompt
'    DoEvents
'    Close #FileNum

End Function

'******************************************************************
'関数名：GetStart
'引数　：nRootPath   検索を開始する最上層のディレクトリへのパス
'　　　：nPattern    パス内のファイル名と対応させる検索文字列
'戻り値：なし
'機能　：指定されたフォルダーに格納されている全てのファイルのパスを配列化
'******************************************************************
Public Function GetFilesMostDeep(ByVal nRootPath As String, ByVal nPattern As String) As String()
    
    Dim lCopy         As Long
    Dim lIndex        As Long
    Dim lBounds       As Long
    Dim nReturns()    As String
    Dim nFilePathes() As String
    Dim hRootFolder   As Folder
    Dim hSubFolder    As Folder
    Dim hFile         As File

    ' 0件の配列を作成
    ReDim nReturns(0)

    ' 検索文字列をすべて大文字にする (Like 演算子が大文字・小文字を区別するため)
    nPattern = StrConv(nPattern, vbUpperCase)

    ' FileSystemObject の新しいインスタンスを生成する
    Dim Fso As FileSystemObject
    Set Fso = New FileSystemObject

    ' Folder オブジェクトを取得する
    Set hRootFolder = Fso.GetFolder(nRootPath)

    ' このディレクトリ内のすべてのファイルを検索する
    For Each hFile In hRootFolder.Files
        If StrConv(hFile.Name, vbUpperCase) Like nPattern Then
            lIndex = lIndex + 1
            ReDim Preserve nReturns(lIndex)
            nReturns(lIndex) = hFile.Path
        End If
    Next hFile

    ' このディレクトリ内のすべてのサブディレクトリを検索する (再帰)
    For Each hSubFolder In hRootFolder.SubFolders
        nFilePathes() = GetFilesMostDeep(hSubFolder.Path, nPattern)

        ' ファイルが格納されている要素数を取得する
        lBounds = UBound(nFilePathes())

        ' 要素数が 1 以上ならば再帰元の配列へコピーする
        If lBounds >= 1 Then
            ReDim Preserve nReturns(lIndex + lBounds)

            For lCopy = 1 To lBounds
                nReturns(lIndex + lCopy) = nFilePathes(lCopy)
            Next lCopy

            lIndex = lIndex + lBounds
        End If
    Next hSubFolder

    ' 不要になった時点で破棄する
    Set Fso = Nothing
    Set hFile = Nothing
    Set hSubFolder = Nothing
    Set hRootFolder = Nothing

    ' 取得したすべてのファイルを返す
    GetFilesMostDeep = nReturns
End Function




