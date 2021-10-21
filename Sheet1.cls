VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' 本プログラムは、Microsoft Scripting Runtimeを有効にして使用する

' フォルダパスと検索するファイル名（ワイルドカード使用可）を定義
Const SEARCH_FOLDER As String = "C:\Users\kazuo\OneDrive\ドキュメント\Excelマクロ\test\"
Const FILE_SEARCH_PATTERN As String = "*_test.xlsx"


' データ収集時の検索のキーとなる情報を格納するための構造体
Private Type ScrapingKey
    sheetName As String
    cellTitle As String
    offsetX As Long
    offsetY As Long
End Type


' 処理本体
' Excelファイル内の表構成やデータ収集の処理方法を変える場合、修正が必要
Public Sub ScrapingExcelFileData()
    
    Dim thisWS As Worksheet
    Set thisWS = ThisWorkbook.ActiveSheet
    Dim targetWB As Workbook
    
    Dim row As Long
    Dim col As Long
    Dim i As Long
    
    Dim files: files = GetFilesForAllDirectories(SEARCH_FOLDER, FILE_SEARCH_PATTERN)
    Dim file
    
    ' データ収集時の検索のキーとなる情報を配列に格納
    ' キーとなる情報は、"SheetName:CellTitle:OffsetX:OffsetY"で表現されている前提の処理
    Dim scrapingKeys() As ScrapingKey
    Dim scrapingKeyString As String
    Dim tmp As Variant
    col = 3
    i = 0
    Do While (thisWS.Cells(2, col) <> "")
        ReDim Preserve scrapingKeys(i)
        scrapingKeyString = thisWS.Cells(2, col)
        tmp = Split(scrapingKeyString, ":")
        scrapingKeys(i).sheetName = tmp(0)
        scrapingKeys(i).cellTitle = tmp(1)
        scrapingKeys(i).offsetX = tmp(2)
        scrapingKeys(i).offsetY = tmp(3)
        col = col + 1
        i = i + 1
    Loop
        
    
    ' データ収集結果の出力処理
    Application.ScreenUpdating = False
    Dim beforeCalculation As XlCalculation
    beforeCalculation = Application.Calculation
    Application.Calculation = xlCalculationManual
    
    row = 3
    For Each file In files
        thisWS.Cells(row, 1) = file.Path
        thisWS.Cells(row, 2) = file.DateLastModified
        
        Set targetWB = Workbooks.Open(thisWS.Cells(row, 1), ReadOnly:=True)
            
        col = 3
        For i = 0 To UBound(scrapingKeys)
            thisWS.Cells(row, col + i) = GetDataFromTargetWB(targetWB, scrapingKeys(i).sheetName, scrapingKeys(i).cellTitle, scrapingKeys(i).offsetX, scrapingKeys(i).offsetY)
        Next
                
        targetWB.Close
            
        row = row + 1
    Next
    
    Application.ScreenUpdating = True
    Application.Calculation = beforeCalculation
    
End Sub

Public Function GetDataFromTargetWB(targetWB As Workbook, sheetName As String, cellTitle As String, offsetX As Long, offsetY As Long) As String
    Dim targetWS As Worksheet
    Set targetWS = Nothing
    On Error Resume Next
    Set targetWS = targetWB.Sheets(sheetName)
    If targetWS Is Nothing Then
        MsgBox "シートの検索に失敗しました"
    Else
        Dim foundCell As Range
        Set foundCell = Nothing
        Set foundCell = targetWS.Cells.Find(cellTitle)
        If foundCell Is Nothing Then
            MsgBox "セルタイトルの検索に失敗しました"
        Else
            GetDataFromTargetWB = foundCell.Offset(offsetY, offsetX).Value
        End If
    End If
    
End Function


'@es2さんの作成されたモジュールをベースに正規表現ではなくワイルドカードしか使えない処理に変更
'https://qiita.com/es2/items/eeb3a02891ff50b6dabc
Public Function GetFilesForAllDirectories(currentFolder As String, fileNamePattern As String) As Variant
    Dim filedic As Object
    Set filedic = CreateObject("scripting.dictionary")
    
    Call GetFilePath(currentFolder, fileNamePattern, filedic)
    GetFilesForAllDirectories = filedic.Keys

End Function

Private Sub GetFilePath(folderPath As String, fileNamePattern As String, ByRef filedic As Object)

    Dim fso As Object: Set fso = CreateObject("scripting.Filesystemobject")
    Dim subfolder As Object

    For Each subfolder In fso.GetFolder(folderPath).SubFolders
        Call GetFilePath(subfolder.Path, fileNamePattern, filedic)
    Next

    Dim file 'as Fileとしたいですが for eachで帰ってくる要素はVariant限定なのでas Fileはデバッグ以外で使用できません
    For Each file In fso.GetFolder(folderPath).files
        If LCase(file.name) Like LCase(fileNamePattern) Then
            filedic.add file, 0
        End If
    Next

End Sub

