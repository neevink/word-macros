'''
''' Neevin Kirill P3113
'''
Private Sub TranslateText_Click()
    
    For i = 1 To ActiveDocument.Content.Words.Count 'Перебираем слова в документе
        j = 1
        
        Do While j <= ActiveDocument.Content.Words.Item(i).Characters.Count 'Перебираем символы в каждом слове
            Do ' Аналог Continue :)

                If (ActiveDocument.Content.Words.Item(i).Characters.Item(j) = "а" Or ActiveDocument.Content.Words.Item(i).Characters.Item(j) = "А") Then
                    ActiveDocument.Content.Words.Item(i).Characters.Item(j).InsertAfter ("ка")
                    j = j + 3
                    Exit Do
                End If

                If (ActiveDocument.Content.Words.Item(i).Characters.Item(j) = "е" Or ActiveDocument.Content.Words.Item(i).Characters.Item(j) = "Е") Then
                    ActiveDocument.Content.Words.Item(i).Characters.Item(j).InsertAfter ("ке")
                    j = j + 3
                    Exit Do
                End If

                If (ActiveDocument.Content.Words.Item(i).Characters.Item(j) = "ё" Or ActiveDocument.Content.Words.Item(i).Characters.Item(j) = "Е") Then
                    ActiveDocument.Content.Words.Item(i).Characters.Item(j).InsertAfter ("кё")
                    j = j + 3
                    Exit Do
                End If

                If (ActiveDocument.Content.Words.Item(i).Characters.Item(j) = "и" Or ActiveDocument.Content.Words.Item(i).Characters.Item(j) = "И") Then
                    ActiveDocument.Content.Words.Item(i).Characters.Item(j).InsertAfter ("ки")
                    j = j + 3
                    Exit Do
                End If

                If (ActiveDocument.Content.Words.Item(i).Characters.Item(j) = "о" Or ActiveDocument.Content.Words.Item(i).Characters.Item(j) = "О") Then
                    ActiveDocument.Content.Words.Item(i).Characters.Item(j).InsertAfter ("ко")
                    j = j + 3
                    Exit Do
                End If

                If (ActiveDocument.Content.Words.Item(i).Characters.Item(j) = "у" Or ActiveDocument.Content.Words.Item(i).Characters.Item(j) = "У") Then
                    ActiveDocument.Content.Words.Item(i).Characters.Item(j).InsertAfter ("ку")
                    j = j + 3
                    Exit Do
                End If

                If (ActiveDocument.Content.Words.Item(i).Characters.Item(j) = "и" Or ActiveDocument.Content.Words.Item(i).Characters.Item(j) = "И") Then
                    ActiveDocument.Content.Words.Item(i).Characters.Item(j).InsertAfter ("ки")
                    j = j + 3
                    Exit Do
                End If

                If (ActiveDocument.Content.Words.Item(i).Characters.Item(j) = "э" Or ActiveDocument.Content.Words.Item(i).Characters.Item(j) = "Э") Then
                    ActiveDocument.Content.Words.Item(i).Characters.Item(j).InsertAfter ("кэ")
                    j = j + 3
                    Exit Do
                End If

                If (ActiveDocument.Content.Words.Item(i).Characters.Item(j) = "ю" Or ActiveDocument.Content.Words.Item(i).Characters.Item(j) = "Ю") Then
                    ActiveDocument.Content.Words.Item(i).Characters.Item(j).InsertAfter ("кю")
                    j = j + 3
                    Exit Do
                End If

                If (ActiveDocument.Content.Words.Item(i).Characters.Item(j) = "я" Or ActiveDocument.Content.Words.Item(i).Characters.Item(j) = "Я") Then
                    ActiveDocument.Content.Words.Item(i).Characters.Item(j).InsertAfter ("кя")
                    j = j + 3
                    Exit Do
                End If

                j = j + 1
            Loop While False
        Loop
    Next i
End Sub