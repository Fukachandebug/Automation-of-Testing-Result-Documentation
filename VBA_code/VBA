Dim Evi_Bk As Workbook
Dim Evi_Sh As Worksheet
Dim ws As Worksheet
Dim destRow As Long
Dim lastRow As Long
Dim rowNum As Long
Dim obj As Object
Dim GroupedObjects As Object
Dim ratio As Double
Dim Manual_Sh As Worksheet
Dim stepNumber As String


'☆手順番号を転記 (シート分けない)☆


'☆エビデンスシートに転記及び貼付け☆

Sub エビシートに手順番号及びエビデンスIDを転記()

    'エビデンスシートを開く
    Set Evi_Bk = Workbooks.Open("C:\Users\81908\OneDrive\デスクトップ\マクロ試作\エビデンスシート.xlsx")

    'エビデンスシートを取得
    Set Evi_Sh = Evi_Bk.Sheets("Sheet1")

    'このブックのシートを指定
    Set ws = ThisWorkbook.Sheets("Sheet1")

    'このブックの最終行を取得
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    '手順番号を転記する開始行
    destRow = 4
    
    'エビデンスシートへの転記
    For rowNum = 4 To lastRow Step 35
        
        Evi_Sh.Range("A" & destRow).Value = ws.Range("A" & rowNum).Value
                                    
        If destRow = 4 Then
        
            'エビデンスIDも転記
            Evi_Sh.Range("B" & destRow + 1).Value = ws.Range("B" & rowNum + 1).Value + ws.Range("D" & rowNum + 1).Value
        
            '罫線引く
            Evi_Sh.Range("A" & destRow & ": AX" & destRow + 39).BorderAround LineStyle:=xlContinuous
        
            '次のエビデンスの開始行
            destRow = destRow + 40
            
        ElseIf destRow <> 4 Then
        
            'エビデンスIDも転記
            Evi_Sh.Range("B" & destRow + 1).Value = ws.Range("B" & rowNum + 1).Value + ws.Range("D" & rowNum + 1).Value
        
            '罫線引く
            Evi_Sh.Range("A" & destRow & ": AX" & destRow + 44).BorderAround LineStyle:=xlContinuous
        
            '次のエビデンスの開始行
            destRow = destRow + 45
            
        End If
                            
    Next rowNum
                       
End Sub


' オブジェクトのサイズをAX列までに収める
Sub ResizeObjectToAX(obj As Object)
    Dim maxColumn As String
    maxColumn = "AX"
    
    ' オブジェクトの幅を調整
    If obj.Left + obj.Width > Evi_Sh.Columns(maxColumn).Left Then
        obj.Width = Evi_Sh.Columns(maxColumn).Left - obj.Left
    End If
    
    ' オブジェクトの高さを調整
    If obj.Top + obj.Height > Evi_Sh.Rows(destRow + 30).Top Then
        obj.Height = Evi_Sh.Rows(destRow + 30).Top - obj.Top
    End If
End Sub


Sub TransferObjects(ws As Worksheet, Evi_Sh As Worksheet)
    ' エビデンスシートの開始行
    destRow = 6

    ' このブックの最終行を取得
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' 各範囲のオブジェクトをエビデンスシートにまとめて貼り付け
    For rowNum = 4 To lastRow Step 35
        ' 同じ枠内のオブジェクトを取得
        Dim objArray() As Shape
        Dim objCount As Integer
        objCount = 0

        ' シート内の全てのShapeを確認
        For Each shp In ws.Shapes
            If Not Intersect(shp.TopLeftCell, ws.Range("A" & rowNum & ":AX" & (rowNum + 34))) Is Nothing Then
                ' 各オブジェクトを一時的な配列に保存
                ReDim Preserve objArray(objCount)
                Set objArray(objCount) = shp
                objCount = objCount + 1
            End If
        Next shp

        If objCount > 0 Then
            If objCount > 1 Then
                ' 複数のオブジェクトがある場合、グループを作成
                Dim groupedShapes As Shape
                Dim arrNames() As String
                ReDim arrNames(0 To objCount - 1)

                ' グループするオブジェクトの名前を配列に格納
                For i = 0 To objCount - 1
                    arrNames(i) = objArray(i).Name
                Next i

                Set groupedShapes = ws.Shapes.Range(arrNames).group

                ' グループをエビデンスシートに貼り付け
                groupedShapes.Copy
            Else
                ' オブジェクトが1つの場合、単体でコピー
                objArray(0).Copy
            End If

            ' クリップボードの内容をエビデンスシートに貼り付け
            Evi_Sh.Activate
            Evi_Sh.Cells(destRow, 2).PasteSpecial
            Application.CutCopyMode = False  ' クリップボードをクリア

            ' グループの位置を再設定
            Evi_Sh.Shapes(Evi_Sh.Shapes.Count).Top = Evi_Sh.Cells(destRow, 2).Top
            Evi_Sh.Shapes(Evi_Sh.Shapes.Count).Left = Evi_Sh.Cells(destRow, 2).Left

            ' グループのサイズを相対的に変更
            Dim ratio As Double
            ratio = Evi_Sh.Range("AU1").Width / objArray(0).Width

            Evi_Sh.Shapes(Evi_Sh.Shapes.Count).LockAspectRatio = msoTrue
            Evi_Sh.Shapes(Evi_Sh.Shapes.Count).Width = objArray(0).Width * ratio * 40
            Evi_Sh.Shapes(Evi_Sh.Shapes.Count).Height = objArray(0).Height * ratio * 40

            ' 次の基点を指定の間隔でずらす
            If destRow = 6 Then
                ' 2枚目のオブジェクトは40行間隔で貼り付ける
                destRow = destRow + 40
            ElseIf destRow <> 6 Then
                ' 3枚目のオブジェクトは45行間隔で貼り付ける
                destRow = destRow + 45
            End If
        End If
    Next rowNum
End Sub


'☆オブジェクト貼付け（グルーピング処理なし）☆'
Sub エビシートにオブジェクト貼付け()

    ' エビデンスシートの開始行
    destRow = 6

    ' エビデンスシートを開く
    Set Evi_Bk = Workbooks.Open("C:\Users\fuu_m\OneDrive\デスクトップ\マクロ試作\エビデンスシート.xlsx")

    ' エビデンスシートを取得
    Set Evi_Sh = Evi_Bk.Sheets(1)

    ' このブックのシートを指定
    Set ws = ThisWorkbook.Sheets(1)

    ' このブックの最終行を取得
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' 各範囲のオブジェクトをエビデンスシートに貼り付け
    For rowNum = 4 To lastRow Step 35
    
        ' 結合セル内のすべてのShapeを取得
        For Each obj In ws.Shapes
        
            ' Shapeが指定した範囲内にあるか確認
            If Not Intersect(obj.TopLeftCell, ws.Range("A" & rowNum & ":AX" & (rowNum + 34))) Is Nothing Then
            
                ' 範囲内のShapeが一つの場合はそのまま貼り付け
                obj.Copy
                Evi_Sh.Cells(destRow, 2).PasteSpecial
                ' サイズをAX列までに収める
                ResizeObjectToAX Evi_Sh.Shapes(Evi_Sh.Shapes.Count)
                ' 位置の調整なども必要であれば追加する
                ' ...
                Application.CutCopyMode = False
                
            End If
            
        Next obj
    
        ' 次の基点を指定の間隔でずらす
        If destRow = 6 Then
            
            ' 2枚目のオブジェクトは40行間隔で貼り付ける
            destRow = destRow + 40
            
        ElseIf destRow <> 6 Then
            
            ' 3枚目のオブジェクトは45行間隔で貼り付ける
            destRow = destRow + 45
                
        End If
        
    Next rowNum

End Sub


'☆完成版（赤枠の図形を表示）☆
Sub エビシートにオブジェクト貼付け3()
    ' エビデンスシートの開始行
    destRow = 4

    ' エビデンスシートを開く
    Set Evi_Bk = Workbooks.Open("C:\Users\fuu_m\OneDrive\デスクトップ\マクロ試作\エビデンスシート.xlsx")

    ' エビデンスシートを取得
    Set Evi_Sh = Evi_Bk.Sheets(1)

    ' このブックのシートを指定
    Set ws = ThisWorkbook.Sheets(1)

    ' このブックの最終行を取得
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' 各範囲のオブジェクトをエビデンスシートに貼り付け
    For rowNum = 4 To lastRow Step 35
    
        ' 範囲内の手順番号を取得
        Dim stepNumber As String
        stepNumber = ws.Range("A" & rowNum).Text


        ' シート内の全てのShapeを確認
        Dim shapeFound As Boolean
        shapeFound = False
        
        For Each shp In ws.Shapes
            ' Shapeが指定した範囲内にあるか確認
            If Not Intersect(shp.TopLeftCell, ws.Range("A" & rowNum & ":AX" & (rowNum + 34))) Is Nothing Then
                ' オブジェクトをコピー
                shp.Copy
                
                ' 対応する手順番号の下に貼り付け
                With Evi_Sh
                    .Cells(destRow, 1).Value = stepNumber
                    .Cells(destRow + 2, 2).PasteSpecial
                                        
                    ' グループの位置を再設定
                    .Shapes(Evi_Sh.Shapes.Count).Top = .Cells(destRow + 2, 2).Top
                    .Shapes(Evi_Sh.Shapes.Count).Left = .Cells(destRow + 2, 2).Left
            
                    ' グループのサイズを相対的に変更
                    Dim ratio As Double
                    ratio = .Range("AU1").Width / shp.Width
                    .Shapes(Evi_Sh.Shapes.Count).LockAspectRatio = msoTrue
                    .Shapes(Evi_Sh.Shapes.Count).Width = shp.Width * ratio * 35
                    .Shapes(Evi_Sh.Shapes.Count).Height = shp.Height * ratio * 35
                                                            
                    ' 貼り付けたオブジェクトを最背面に移動
                    Evi_Sh.Shapes(Evi_Sh.Shapes.Count).ZOrder msoSendToBack
                                                            
                    ' 罫線引く
                    .Range("A" & destRow & ": AX" & destRow + 44).BorderAround LineStyle:=xlContinuous
                    
                End With
                
                                
                shapeFound = True
                
                With Evi_Sh.Shapes.AddShape(msoShapeRectangle, Evi_Sh.Cells(destRow + 1, 45).Left, Evi_Sh.Cells(destRow + 1, 45).Top, 50, 50)
                
                    ' 枠線の色を赤に設定
                    .Line.ForeColor.RGB = RGB(255, 0, 0)
                    
                    ' 塗りつぶしなし
                    .Fill.Visible = msoFalse
                    
                    ' 枠線を太くする
                    .Line.Weight = 3
                    
                    ' サイズ変更
                    .LockAspectRatio = msoFalse
                    .Width = 200
                    .Height = 50
                    
                    ' 図形にテキストを追加
                    .TextFrame.Characters.Text = "確認した"
                    
                    ' テキストの色を赤に設定
                    .TextFrame.Characters.Font.Color = RGB(255, 0, 0)
                    
                    ' テキストのフォントサイズやスタイルを設定する場合
                    .TextFrame.Characters.Font.Size = 12
                    
                End With
                
                    ' 次のエビデンスの開始行
                    destRow = destRow + 45
                    
            End If
            
        Next shp
        
        ' Shapeが範囲内に存在しない場合、手順番号だけを貼り付け
        If Not shapeFound Then
            Evi_Sh.Cells(destRow, 1).Value = stepNumber
            
            ' 罫線引く
            Evi_Sh.Range("A" & destRow & ": AX" & destRow + 44).BorderAround LineStyle:=xlContinuous
            destRow = destRow + 45
        End If
    Next rowNum
End Sub


Function FindRowByStepNumber(ByRef stepNumber As Variant, ByRef ws As Worksheet) As Long
    ' 手順書のA列から指定された手順番号の行を探す関数
    Dim foundCell As Range
    Set foundCell = ws.Range("A:A").Find(What:=stepNumber, LookIn:=xlValues, LookAt:=xlWhole)

    If Not foundCell Is Nothing Then
        FindRowByStepNumber = foundCell.Row
    Else
        ' 見つからない場合は0を返す
        FindRowByStepNumber = 0
    End If
End Function

'☆コメント表示可能だが、オブジェクトの貼付け位置がズレる☆'
Sub エビシートにオブジェクト貼付け5()
    ' エビデンスシートの開始行
    destRow = 4

    ' エビデンスシートを開く
    Set Evi_Bk = Workbooks.Open("C:\Users\fuu_m\OneDrive\デスクトップ\マクロ試作\エビデンスシート.xlsx")

    ' エビデンスシートを取得
    Set Evi_Sh = Evi_Bk.Sheets(1)

    ' このブックのシートを指定
    Set ws = ThisWorkbook.Sheets(1)
    
    ' 手順書を開く
    Set Manual_Bk = Workbooks.Open("C:\Users\fuu_m\OneDrive\デスクトップ\マクロ試作\手順書.xlsx")

    ' 手順書のシートを取得
    Set Manual_Sh = Manual_Bk.Sheets(1)

    ' このブックの最終行を取得
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' 各範囲のオブジェクトをエビデンスシートに貼り付け
    For rowNum = 4 To lastRow Step 35
    
        ' 範囲内の手順番号を取得
        stepNumber = ws.Range("A" & rowNum).Value

        ' シート内の全てのShapeを確認
        Dim shapeFound As Boolean
        shapeFound = False
        
        For Each shp In ws.Shapes
            ' Shapeが指定した範囲内にあるか確認
            If Not Intersect(shp.TopLeftCell, ws.Range("A" & rowNum & ":AX" & (rowNum + 34))) Is Nothing Then
            
                            ' デバッグ情報を追加
                Debug.Print "Step Number: " & stepNumber
                Debug.Print "Copied Object Top: " & shp.Top
                Debug.Print "Copied Object Left: " & shp.Left
                Debug.Print "Destination Top: " & Evi_Sh.Cells(destRow + 2, 2).Top
                Debug.Print "Destination Left: " & Evi_Sh.Cells(destRow + 2, 2).Left

                ' オブジェクトをコピー
                shp.Copy
                
                
                ' 対応する手順番号の下に貼り付け
                With Evi_Sh
                    .Cells(destRow, 1).Value = stepNumber
                    .Cells(destRow + 3, 2).PasteSpecial
                                        
                    ' グループの位置を再設定
                    .Shapes(Evi_Sh.Shapes.Count).Top = .Cells(destRow + 2, 2).Top
                    .Shapes(Evi_Sh.Shapes.Count).Left = .Cells(destRow + 2, 2).Left
            
                    ' グループのサイズを相対的に変更
                    Dim ratio As Double
                    ratio = .Range("AU1").Width / shp.Width
                    .Shapes(Evi_Sh.Shapes.Count).LockAspectRatio = msoTrue
                    .Shapes(Evi_Sh.Shapes.Count).Width = shp.Width * ratio * 35
                    .Shapes(Evi_Sh.Shapes.Count).Height = shp.Height * ratio * 35
                                                            
                    ' 貼り付けたオブジェクトを最背面に移動
                    Evi_Sh.Shapes(Evi_Sh.Shapes.Count).ZOrder msoSendToBack
                                                            
                    ' 罫線引く
                    .Range("A" & destRow & ": AX" & destRow + 44).BorderAround LineStyle:=xlContinuous
                    
                End With
                
                ' 手順書から対応するセルの値をコピーして貼り付け
                Dim foundRow As Long
                foundRow = FindRowByStepNumber(stepNumber, Manual_Sh)

                If foundRow <> 0 Then

                    ' 赤枠の図形にテキストを表示
                    With Evi_Sh.Shapes.AddShape(msoShapeRectangle, Evi_Sh.Cells(destRow + 1, 45).Left, Evi_Sh.Cells(destRow + 1, 45).Top, 50, 50)
                        .Line.ForeColor.RGB = RGB(255, 0, 0)
                        .Fill.Visible = msoFalse
                        .Line.Weight = 3
                        .LockAspectRatio = msoFalse
                        .Width = 200
                        .Height = 50
                        .TextFrame.Characters.Text = Manual_Sh.Cells(foundRow, 4).Value
                        .TextFrame.Characters.Font.Color = RGB(255, 0, 0)
                        .TextFrame.Characters.Font.Size = 12
                    End With
                End If
    
    
                shapeFound = True
                
                ' 次のエビデンスの開始行
                destRow = destRow + 45
            End If
        Next shp
        
        ' Shapeが範囲内に存在しない場合、手順番号だけを貼り付け
        If Not shapeFound Then
            Evi_Sh.Cells(destRow, 1).Value = stepNumber
            
            ' 罫線引く
            Evi_Sh.Range("A" & destRow & ": AX" & destRow + 44).BorderAround LineStyle:=xlContinuous
            
            destRow = destRow + 45
        End If
    Next rowNum
End Sub


