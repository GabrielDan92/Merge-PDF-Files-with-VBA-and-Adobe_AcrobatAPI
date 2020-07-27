Public Sub PDFMerge(ByVal month As String, ByVal day As String)

Dim fso As Object, f As Object, sf As Object, sf2 As Object, sf3 As Object
Dim pdfFile As Object, finalPDFName As String, originalFolderPath As String
Dim strPDFs() As String

answer = MsgBox("Is the path in cell K2 correct?", vbQuestion + vbYesNo + vbDefaultButton2)
If answer <> vbYes Then Exit Sub

stringMonth = month
Select Case month
    Case "Ianuarie"
        month = 1
    Case "Februarie"
        month = 2
    Case "Martie"
        month = 3
    Case "Aprilie"
        month = 4
    Case "Mai"
        month = 5
    Case "Iunie"
        month = 6
    Case "Iulie"
        month = 7
    Case "August"
        month = 8
    Case "Septembrie"
        month = 9
    Case "Octombrie"
        month = 10
    Case "Noiembrie"
        month = 11
    Case "Decembrie"
        month = 12
End Select
dateConcat = day & " " & month & ", " & Year(Now())
'check the date and convert it if it's valid
On Error Resume Next
dateToExcel = DateValue(dateConcat)
On Error GoTo 0
If dateToExcel = "" Then
    MsgBox "Wrong date!"
    Exit Sub
End If
                
originalFolderPath = Sheet1.Range("K2").Value
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.GetFolder(originalFolderPath)
Set dict = New Dictionary
dict.CompareMode = vbTextCompare

customYear = "2020"
PDFName = "Merged_PDF " & day & "." & month & "." & customYear
finalPDFName = PDFName
EmailSubject = finalPDFName
    
For Each sf In f.subfolders                                                                 'month subfolder - starting folder
    If InStr(UCase(sf.Name), UCase(stringMonth)) > 0 Then
        Debug.Print "Accessing month folder " & sf.Name
        For Each sf2 In sf.subfolders                                                       'day subfolder
            If Int(Left(sf2.Name, 2)) = Int(day) Then
                Debug.Print "Accessing day folder " & sf2.Name
                For Each sf3 In sf2.subfolders                                              '3rd level subfolder - final folder
                    Debug.Print "Accessing 3rd level subfolder " & sf3.Name
                    FolderName = Trim(sf3.Name)
                    For Each pdfFile In sf3.Files                                           'target PDF files and send each one to a dict
                        If fso.getextensionName(pdfFile.Path) = "pdf" Then
                            Debug.Print pdfFile.Name
                            dict(pdfFile.Path) = pdfFile.Path
                        End If
                    Next
                    If dict.Count > 0 Then
                        ReDim strPDFs(dict.Count - 1)
                        i = 0
                        For Each Key In dict                                                'all the dict files are pushed to a string array
                            strPDFs(i) = dict(Key)                                          'then merged using Adobe
                            i = i + 1
                        Next
                        finalPDFName = sf3.Path + "\" + finalPDFName + ".pdf"
                        bSuccess = MergePDFs(strPDFs, finalPDFName)                         'call the merge function that uses Adobe's API
                        dict.RemoveAll
                        finalPDFName = PDFName
                    End If
                Next
            End If
        Next
    End If
Next

MsgBox "Done!"

End Sub

'=======================================================================================
'the PDF function was not written by me
'credit to: https://wellsr.com/vba/2017/word/combine-pdfs-with-vba-and-adobe-acrobat/

Private Function MergePDFs(arrFiles() As String, strSaveAs As String) As Boolean
 
    Dim objCAcroPDDocDestination As Acrobat.CAcroPDDoc
    Dim objCAcroPDDocSource As Acrobat.CAcroPDDoc
    Dim i As Integer
    Dim iFailed As Integer
     
    On Error GoTo NoAcrobat:
    'Initialize the Acrobat objects
    Set objCAcroPDDocDestination = CreateObject("AcroExch.PDDoc")
    Set objCAcroPDDocSource = CreateObject("AcroExch.PDDoc")
     
    'Open Destination, all other documents will be added to this and saved with
    'a new filename
    objCAcroPDDocDestination.Open (arrFiles(LBound(arrFiles))) 'open the first file
     
    'Open each subsequent PDF that you want to add to the original
      'Open the source document that will be added to the destination
        For i = LBound(arrFiles) + 1 To UBound(arrFiles)
            objCAcroPDDocSource.Open (arrFiles(i))
            If objCAcroPDDocDestination.InsertPages(objCAcroPDDocDestination.GetNumPages - 1, objCAcroPDDocSource, 0, objCAcroPDDocSource.GetNumPages, 0) Then
              MergePDFs = True
            Else
              'failed to merge one of the PDFs
              iFailed = iFailed + 1
            End If
            objCAcroPDDocSource.Close
        Next i
    objCAcroPDDocDestination.Save 1, strSaveAs 'Save it as a new name
    objCAcroPDDocDestination.Close
    Set objCAcroPDDocSource = Nothing
    Set objCAcroPDDocDestination = Nothing
     
NoAcrobat:
    If iFailed <> 0 Then
        MergePDFs = False
    End If
    On Error GoTo 0
    
End Function
