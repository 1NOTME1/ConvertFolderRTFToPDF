Sub ConvertFolderRTFToPDF()
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim srcFolderPath As String
    Dim destFolderPath As String
    Dim filePath As String
    Dim savePath As String
    Dim fileName As String
    
    ' Ścieżka do folderu z plikami RTF
    srcFolderPath = "xxx\x\"
    
    ' Ścieżka do folderu, gdzie pliki PDF zostaną zapisane
    destFolderPath = "xxx\x\"
    
    ' Utwórz nową instancję aplikacji Word
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False
    
    ' Ustaw folder źródłowy jako bieżący katalog
    ChDir srcFolderPath
    
    ' Szukaj pierwszego pliku RTF
    fileName = Dir("*.rtf")
    
    Do While fileName <> ""
        ' Pełna ścieżka do pliku RTF
        filePath = srcFolderPath & fileName
        
        ' Pełna ścieżka do pliku PDF
        savePath = destFolderPath & Left(fileName, Len(fileName) - 4) & ".pdf"
        
        ' Otwórz dokument RTF
        Set wordDoc = wordApp.Documents.Open(filePath)
        
        ' Zapisz jako PDF
        wordDoc.SaveAs2 savePath, 17 ' 17 oznacza format PDF
        
        ' Zamknij dokument
        wordDoc.Close
        
        ' Szukaj kolejnego pliku RTF
        fileName = Dir
    Loop
    
    ' Zakończ aplikację Word
    wordApp.Quit
    
    ' Czyszczenie obiektów
    Set wordDoc = Nothing
    Set wordApp = Nothing

    MsgBox "Konwersja folderu zakończona pomyślnie"
End Sub

