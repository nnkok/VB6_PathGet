Attribute VB_Name = "Mdl_GetPath"
Option Explicit

Public stfilePath() As String   '�擾�����t�@�C���p�X��S�Ĕz��

'******************************************************************
'�֐����FGetStart
'�����@�FstFolderPath
'�߂�l�F�Ȃ�
'�@�\�@�F�w�肳�ꂽ�t�H���_�[�Ɋi�[����Ă���S�Ẵt�@�C���̃p�X��z��
'******************************************************************
Public Function GetStart(ByVal stFolderPath As String) As Boolean

    Dim lIndex          As Long
    Dim nPrompt         As String
    Dim nFilePath       As String
    Dim nFilePathes()   As String
'    Dim TextPath        As String
'    Dim FileNum         As Variant
    Dim lCont           As Long

    '����"*.*"���w�肷�邱�ƂőS�Ẵt�@�C�����Ώ�(�ύX���邱�ƂŊg���q�̎w��Ȃǉ\�@��j"*.txt"�@��TEXT�t�@�C���̌���)
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
'�֐����FGetStart
'�����@�FnRootPath   �������J�n����ŏ�w�̃f�B���N�g���ւ̃p�X
'�@�@�@�FnPattern    �p�X���̃t�@�C�����ƑΉ������錟��������
'�߂�l�F�Ȃ�
'�@�\�@�F�w�肳�ꂽ�t�H���_�[�Ɋi�[����Ă���S�Ẵt�@�C���̃p�X��z��
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

    ' 0���̔z����쐬
    ReDim nReturns(0)

    ' ��������������ׂđ啶���ɂ��� (Like ���Z�q���啶���E����������ʂ��邽��)
    nPattern = StrConv(nPattern, vbUpperCase)

    ' FileSystemObject �̐V�����C���X�^���X�𐶐�����
    Dim Fso As FileSystemObject
    Set Fso = New FileSystemObject

    ' Folder �I�u�W�F�N�g���擾����
    Set hRootFolder = Fso.GetFolder(nRootPath)

    ' ���̃f�B���N�g�����̂��ׂẴt�@�C������������
    For Each hFile In hRootFolder.Files
        If StrConv(hFile.Name, vbUpperCase) Like nPattern Then
            lIndex = lIndex + 1
            ReDim Preserve nReturns(lIndex)
            nReturns(lIndex) = hFile.Path
        End If
    Next hFile

    ' ���̃f�B���N�g�����̂��ׂẴT�u�f�B���N�g������������ (�ċA)
    For Each hSubFolder In hRootFolder.SubFolders
        nFilePathes() = GetFilesMostDeep(hSubFolder.Path, nPattern)

        ' �t�@�C�����i�[����Ă���v�f�����擾����
        lBounds = UBound(nFilePathes())

        ' �v�f���� 1 �ȏ�Ȃ�΍ċA���̔z��փR�s�[����
        If lBounds >= 1 Then
            ReDim Preserve nReturns(lIndex + lBounds)

            For lCopy = 1 To lBounds
                nReturns(lIndex + lCopy) = nFilePathes(lCopy)
            Next lCopy

            lIndex = lIndex + lBounds
        End If
    Next hSubFolder

    ' �s�v�ɂȂ������_�Ŕj������
    Set Fso = Nothing
    Set hFile = Nothing
    Set hSubFolder = Nothing
    Set hRootFolder = Nothing

    ' �擾�������ׂẴt�@�C����Ԃ�
    GetFilesMostDeep = nReturns
End Function




