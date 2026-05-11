Attribute VB_Name = "Module1"
Option Explicit

'******** "Calc" button ********
Sub GeneratePrimesAndGaps()
    ' --- 変数定義 ---
    Dim isPrime() As Boolean     ' エラトステネスの篩用配列
    Dim result() As Variant      ' 書き出しデータ用配列 (行, 列)
    Dim i As Long, j As Long
    Dim pCount As Long           ' 素数の個数カウント
    Dim lastPrime As Long        ' 1つ前の素数を保持
    Dim diff As Long             ' 素数間の差
    Dim digitCount As Integer    ' 桁数
    
    Const NDIFF As Long = 20       '素数砂漠

    ' --- 設定 ---
    'Const MAX_VAL As Long = 1000000 ' 100万まで
    Dim MAX_VAL As Long
    MAX_VAL = Range("G1") ' MAX 16290000 ?
    
    ' 画面更新と自動計算を停止（高速化のため）
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' --- 1. エラトステネスの篩による高速素数判定 ---
    ReDim isPrime(MAX_VAL)
    For i = 2 To MAX_VAL: isPrime(i) = True: Next i

    ' 2から√MAX_VALまでループして倍数を消す
    For i = 2 To Sqr(MAX_VAL)
        If isPrime(i) Then
            For j = i * i To MAX_VAL Step i
                isPrime(j) = False
            Next j
        End If
        ' ループの節目で「応答なし」を回避
        If i Mod 100 = 0 Then DoEvents
    Next i

    ' 素数の総数をカウント（配列のサイズ決定用）
    pCount = 0
    For i = 2 To MAX_VAL
        If isPrime(i) Then pCount = pCount + 1
    Next i

    ' 結果格納用配列を準備 (行数: pCount, 列数: 4)
    ReDim result(1 To pCount, 1 To 4)

    ' --- 2. 条件に基づき配列へデータを格納 ---
    pCount = 0
    lastPrime = 0

    For i = 2 To MAX_VAL
        If isPrime(i) Then
            pCount = pCount + 1
            result(pCount, 1) = i ' A列: 素数

            ' B列: 1つ前の素数との差
            If lastPrime = 0 Then
                result(pCount, 2) = "-" ' 最初の素数(2)には差がない
            Else
                diff = i - lastPrime
                result(pCount, 2) = diff

                ' C列: 差が2（双子素数）なら「*」
                If diff = 2 Then
                    result(pCount, 3) = "*"
                End If

                ' D列: 差が NDIFF以上なら桁数分「*」    '100
                If diff >= NDIFF Then
                    digitCount = diff / 10
                    result(pCount, 4) = String(digitCount, "*")
                End If
            End If

            lastPrime = i ' 今回の素数を「1つ前」として保存
        End If

        ' 定期的に制御を戻す（巨大ループのフリーズ回避）
        If i Mod 10000 = 0 Then DoEvents
    Next i

    ' --- 3. シートへの一括書き出し ---
    With ThisWorkbook.Sheets(1)
        '.Cells.ClearContents ' 前のデータをクリア
        If ActiveSheet.FilterMode = True Then ActiveSheet.ShowAllData   'オートフィルタを解除
        Range("A2:D1048576").Delete 'データを削除
        ' ヘッダーの作成
        .Range("A1:D1").Value = Array("素数", "差", "双子", "素数砂漠" & vbCrLf & NDIFF & "以上")

        ' 配列を一気に流し込む（これが最速）
        .Range("A2").Resize(pCount, 4).Value = result
    End With

    ' 設定を元に戻す
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Range("A2").Select

    MsgBox "処理が完了しました！" & vbCrLf & "素数の数: " & pCount
End Sub

'******** "Clear" button ********
Sub subClear()
    If MsgBox("削除しますか", vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    If ActiveSheet.FilterMode = True Then ActiveSheet.ShowAllData   'オートフィルタを解除
    Range("A2:D1048576").Delete 'データを削除
    Range("G1") = 16290000 '1048576
End Sub


'Sub GeneratePrimesWithSafetySplit()
'    ' ================================================================
'    ' 設定値
'    ' ================================================================
'    ' 16,300,000までの素数は 1,044,424個存在します。
'    ' これはExcelの最大行数 1,048,576行の限界ギリギリの数値です。
'    Const MAX_VAL As Long = 16300000
'    Const BATCH_SIZE As Long = 100000 ' 10万行ごとに分割して転記（エラー回避用）
'
'    ' 変数宣言
'    Dim isPrime() As Boolean     ' エラトステネスの篩用（メモリ節約のためBoolean型）
'    Dim result() As Variant      ' 全計算結果を保持する巨大配列
'    Dim i As Long, j As Long, k As Long
'    Dim pCount As Long           ' 見つかった素数の累計個数
'    Dim lastPrime As Long        ' 差を計算するための1つ前の素数
'    Dim diff As Long             ' 素数間の差
'
'    ' 高速化設定：画面更新と自動計算を停止
'    Application.ScreenUpdating = False
'    Application.Calculation = xlCalculationManual
'
'    ' ================================================================
'    ' 1. エラトステネスの篩（ふるい）による高速判定
'    ' ================================================================
'    ' 指定範囲の全数値を「素数候補(True)」として初期化
'    ReDim isPrime(MAX_VAL)
'    For i = 2 To MAX_VAL: isPrime(i) = True: Next i
'
'    ' 2から√MAX_VALまでループ。既知の素数の倍数を「素数ではない(False)」として消していく
'    ' このアルゴリズムが1000万超えの範囲では最も効率的です。
'    For i = 2 To Sqr(MAX_VAL)
'        If isPrime(i) Then
'            For j = i * i To MAX_VAL Step i
'                isPrime(j) = False
'            Next j
'        End If
'        ' 100ループごとにWindowsに制御を戻し「応答なし」を防止
'        If i Mod 100 = 0 Then DoEvents
'    Next i
'
'    ' 素数の総数をカウントし、結果格納用配列のサイズを確定させる
'    pCount = 0
'    For i = 2 To MAX_VAL
'        If isPrime(i) Then pCount = pCount + 1
'    Next i
'
'    ' 素数の数がExcelの最大行数を超えていないかチェック
'    If pCount > 1048575 Then
'        MsgBox "素数の数がExcelの行数制限を超えました。範囲を小さくしてください。"
'        Exit Sub
'    End If
'
'    ' 結果用配列を再定義 (行, 列)
'    ReDim result(1 To pCount, 1 To 4)
'
'    ' ================================================================
'    ' 2. 各条件の判定と配列への書き込み
'    ' ================================================================
'    pCount = 0
'    lastPrime = 0
'
'    For i = 2 To MAX_VAL
'        If isPrime(i) Then
'            pCount = pCount + 1
'            ' A列: 素数そのもの
'            result(pCount, 1) = i
'
'            If lastPrime <> 0 Then
'                ' B列: 前の素数との差
'                diff = i - lastPrime
'                result(pCount, 2) = diff
'
'                ' C列: 差が2（双子素数）なら「*」を付加
'                If diff = 2 Then result(pCount, 3) = "*"
'
'                ' D列: 差が100以上なら、その素数の桁数分だけ「*」を付加
'                If diff >= 100 Then
'                    result(pCount, 4) = String(Len(CStr(i)), "*")
'                End If
'            Else
'                result(pCount, 2) = "-" ' 最初の素数2用
'            End If
'            lastPrime = i
'        End If
'        ' 1万回ごとにフリーズ回避処理
'        If i Mod 10000 = 0 Then DoEvents
'    Next i
'
'    ' ================================================================
'    ' 3. 分割書き出し（「アプリケーション定義のエラー」対策）
'    ' ================================================================
'    ' Excelの内部制限により、100万要素を超える配列は一度に書き込めないため、
'    ' 小さなブロック（BATCH_SIZE）に切り分けてシートに転記します。
'    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(1)
'    ws.Cells.ClearContents
'    ws.Range("A1:D1").Value = Array("素数", "差", "双子", "100以上差")
'
'    Dim currentBatch() As Variant ' 小分け用の中継配列
'    Dim rowsToCopy As Long        ' 今回書き出す行数
'    Dim startIdx As Long          ' result配列の読み取り開始位置
'
'    For startIdx = 1 To pCount Step BATCH_SIZE
'        ' 残りの行数がBATCH_SIZEより少ない場合の調整
'        If startIdx + BATCH_SIZE - 1 > pCount Then
'            rowsToCopy = pCount - startIdx + 1
'        Else
'            rowsToCopy = BATCH_SIZE
'        End If
'
'        ' 今回のブロック分だけ中継配列を作成
'        ReDim currentBatch(1 To rowsToCopy, 1 To 4)
'        For i = 1 To rowsToCopy
'            For j = 1 To 4
'                currentBatch(i, j) = result(startIdx + i - 1, j)
'            Next j
'        Next i
'
'        ' 指定した開始行（startIdx + 1）からシートに書き込み
'        ws.Cells(startIdx + 1, 1).Resize(rowsToCopy, 4).Value = currentBatch
'        DoEvents
'    Next startIdx
'
'    ' 設定を復元
'    Application.ScreenUpdating = True
'    Application.Calculation = xlCalculationAutomatic
'
'    MsgBox "全素数の抽出が完了しました！" & vbCrLf & "素数個数: " & pCount
'End Sub
