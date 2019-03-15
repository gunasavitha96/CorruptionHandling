Sub ProgressBar()
On Error Resume Next
Dim fd As Office.FileDialog
Dim wdFileName As String
Dim objWord As Object
Dim wdDoc As Object
Dim oldpath As String
Dim newpath As String
Dim tableno As Integer
Dim iTable As Integer
Dim r, c As Integer
Dim count As Integer
Dim iTotalRows, iTotalCols As Integer
Dim RowCount As Integer
Dim row As Integer
Dim myCopy As Document
Dim oRange
Dim oTable
Dim i As Integer
Dim iCols As Integer
Dim rnge As Range
Dim backColor As Long
Dim foreColor As Long
Dim iRows As Integer
Dim rnge1 As Range
Dim rnge2 As String
Dim rnge3 As String
Dim rnge4 As String
Dim position As Integer
Dim iCols1 As Integer
Dim copyfilecount As Integer
Dim sourcefilecount As Integer
Dim myfile As String
count = 0
RowCount = 0
Dim URSTemplate As String
Dim s1 As String
Dim txt As String
Dim txt1 As Range

Dim i1 As Long
Dim i2 As Long
Dim CurrentProgress As Double
Dim ProgressPercentage As Double ' For percentage value in progressbar
Dim BarWidth As Long
Dim ione As Integer
Dim Tcount As Integer
i1 = 0
i2 = 0
ione = 0
Tcount = 0


URSTemplate = "C:\TableSplitting\URS_template.docx"
Dim fsO
Set fsO = CreateObject("Scripting.FileSystemObject")

Set fd = Application.FileDialog(msoFileDialogFilePicker)
With fd
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Word files", "*.docx"
        .Filters.Add "All Files", "*.doc"
        .Filters.Add "All Files", "*.docm"
        .Title = "Select a Word File"
  If .Show = True Then
           wdFileName = Dir(.SelectedItems(1))
End If
End With
If Trim(wdFileName) <> "" Then
Set objWord = CreateObject("Word.Application")
 objWord.Visible = False
 objWord.Activate
 Set wdDoc = GetObject(wdFileName)
 oldpath = fsO.GetParentFolderName(wdDoc.FullName)
 Dim fsoobject
 Set fsoobject = CreateObject("scripting.filesystemobject")
 If fsoobject.FolderExists(oldpath & "\" & "URS~FinalOutput\") Then
 Else
    fsoobject.CreateFolder (oldpath & "\" & "URS~FinalOutput\")
 End If
 newpath = oldpath & "\" & "URS~FinalOutput\"
 
 With wdDoc
   'MsgBox wdDoc.FullName
   
   
   tableno = wdDoc.Tables.count
   If tableno = 0 Then
     MsgBox "This document contains no tables", _
     vbExclamation, "Import Word Table"
   ElseIf tableno > 1 Then
     MsgBox "This Word document contains " & tableno & " tables."
   ElseIf tableno = 1 Then
     MsgBox "There is Only One table in the word."
   End If
 ' Dim fontsize As Integer
   For iTable = 0 To tableno
    With wdDoc.Tables(iTable)
       r = 1
       c = 1
        For r = 1 To 1
            For c = 1 To 1
               If (r = 1) And (wdDoc.Tables(iTable).Cell(r, c).Range.Text Like "*Business*") Then
                     count = count + 1 ' Count the number of tables
               End If
            Next c
        Next r
    End With
   Next iTable
   count = count - 1
   
   
    For iTable = 0 To tableno
    With wdDoc.Tables(iTable)
       r = 1
       c = 1
        For r = 1 To 1
           For c = 1 To 1
             If (r = 1) And (wdDoc.Tables(iTable).Cell(r, c).Range.Text Like "*Business*") Then
                With wdDoc.Tables(iTable)
                iTotalRows = wdDoc.Tables(iTable).rows.count 'To count the total rows
                   RowCount = RowCount + iTotalRows ' To count total rows in the table
               End With
             End If
           Next c
        Next r
    End With
   Next iTable
   RowCount = RowCount - count
   i2 = CLng(RowCount)
   
   'Call InitProgressBar
    
   For iTable = 0 To tableno And (i1 <= i2)
   With wdDoc.Tables(iTable)
       r = 1
       c = 1
        For r = 1 To 1
           For c = 1 To 1
             If (r = 1) And (wdDoc.Tables(iTable).Cell(r, c).Range.Text Like "*Business*") Then
              With wdDoc.Tables(iTable)
                iTotalRows = wdDoc.Tables(iTable).rows.count
                iTotalCols = wdDoc.Tables(iTable).columns.count
                ' MsgBox wdDoc.Tables(iTable).Columns.count
                For row = 2 To iTotalRows
                 If wdDoc.Tables(iTable).rows(row) <> "" Then
                  With wdDoc.Tables(iTable).rows(row)
                  
                  'Set myCopy = objWord.Documents.Add(URSTemplate)
                  
                  Set myCopy = objWord.Documents.Open(filename:=URSTemplate)
                 
                  Set oTable = myCopy.Tables(1)
                  copyfilecount = myCopy.Tables(1).columns.count
                  For iCols = 1 To copyfilecount
                  
                  txt = myCopy.Tables(1).Cell(1, iCols).Range.Text
                  txt = Trim(txt)
                  
                  sourcefilecount = wdDoc.Tables(iTable).rows(row).Cells.count
                 
                    
                    For iCols1 = 1 To sourcefilecount
                    s1 = wdDoc.Tables(iTable).Cell(1, iCols1).Range.Text
                    s1 = Trim(s1)
                   
                    If (StrComp(txt, s1, vbTextCompare) = 0) Then
                  
                       txt1 = wdDoc.Tables(iTable).Cell(row, iCols1).Range.copy
                       backColor = wdDoc.Tables(iTable).Cell(row, iCols).Shading.BackgroundPatternColor
                       foreColor = wdDoc.Tables(iTable).Cell(row, iCols).Shading.ForegroundPatternColor
                       
                        myCopy.Tables(1).Cell(2, iCols).Range.Italic = False
                        
                        
                        
                        myCopy.Tables(1).Cell(2, iCols).Range.Font.ColorIndex = wdBlack
                       
                       myCopy.Tables(1).Cell(2, iCols).Range.Paste
                       myCopy.Tables(1).Cell(2, iCols).Range.Font.Size = myCopy.Tables(1).Cell(1, iCols).Range.Font.Size
                       myCopy.Tables(1).Cell(2, iCols).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
                       
                       oTable.Cell(2, iCols).Shading.BackgroundPatternColor = backColor ' Paste the Backcolor
                       oTable.Cell(2, iCols).Shading.ForegroundPatternColor = foreColor
                       
                       
                     ElseIf StrComp(txt, s1, vbTextCompare) And txt Like "*URS Identifier*" And s1 Like "*Req. Num.*" Then
                       
                       txt1 = wdDoc.Tables(iTable).Cell(row, iCols1).Range.copy
                       backColor = wdDoc.Tables(iTable).Cell(row, iCols).Shading.BackgroundPatternColor
                       foreColor = wdDoc.Tables(iTable).Cell(row, iCols).Shading.ForegroundPatternColor
                       
                        myCopy.Tables(1).Cell(2, iCols).Range.Italic = False
                        myCopy.Tables(1).Cell(2, iCols).Range.Font.ColorIndex = wdBlack
                       
                        
                        
                       
                       myCopy.Tables(1).Cell(2, iCols).Range.Paste
                         myCopy.Tables(1).Cell(2, iCols).Range.Font.Size = myCopy.Tables(1).Cell(1, iCols).Range.Font.Size
                        myCopy.Tables(1).Cell(2, iCols).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
                       oTable.Cell(2, iCols).Shading.BackgroundPatternColor = backColor ' Paste the Backcolor
                       oTable.Cell(2, iCols).Shading.ForegroundPatternColor = foreColor
                      
                     ElseIf StrComp(txt, s1, vbTextCompare) And txt Like "*Release First Implemented*" And s1 Like "*Version Implemented*" Then
                       txt1 = wdDoc.Tables(iTable).Cell(row, iCols1).Range.copy
                       backColor = wdDoc.Tables(iTable).Cell(row, iCols).Shading.BackgroundPatternColor
                       foreColor = wdDoc.Tables(iTable).Cell(row, iCols).Shading.ForegroundPatternColor
                       
                       myCopy.Tables(1).Cell(2, iCols).Range.Italic = False
                       myCopy.Tables(1).Cell(2, iCols).Range.Font.ColorIndex = wdBlack
                      
                       
                        
                       myCopy.Tables(1).Cell(2, iCols).Range.Paste
                       myCopy.Tables(1).Cell(2, iCols).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
                       myCopy.Tables(1).Cell(2, iCols).Range.Font.Size = myCopy.Tables(1).Cell(1, iCols).Range.Font.Size
                       oTable.Cell(2, iCols).Shading.BackgroundPatternColor = backColor ' Paste the Backcolor
                       oTable.Cell(2, iCols).Shading.ForegroundPatternColor = foreColor
                        
                    ElseIf StrComp(txt, s1, vbTextCompare) And txt Like "*Release Last Changed*" And s1 Like "*Version Last Changed*" Then
                       
                       txt1 = wdDoc.Tables(iTable).Cell(row, iCols1).Range.copy
                       backColor = wdDoc.Tables(iTable).Cell(row, iCols).Shading.BackgroundPatternColor
                       foreColor = wdDoc.Tables(iTable).Cell(row, iCols).Shading.ForegroundPatternColor
                       
                        myCopy.Tables(1).Cell(2, iCols).Range.Italic = False
                        myCopy.Tables(1).Cell(2, iCols).Range.Font.ColorIndex = wdBlack
                        
                        
                       myCopy.Tables(1).Cell(2, iCols).Range.Paste
                       myCopy.Tables(1).Cell(2, iCols).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
                        myCopy.Tables(1).Cell(2, iCols).Range.Font.Size = myCopy.Tables(1).Cell(1, iCols).Range.Font.Size
                       oTable.Cell(2, iCols).Shading.BackgroundPatternColor = backColor ' Paste the Backcolor
                       oTable.Cell(2, iCols).Shading.ForegroundPatternColor = foreColor
                      
                        
                       
                    ElseIf StrComp(txt, s1, vbTextCompare) And txt Like "*Risk Determination*" And s1 Like "*Risk Determination*" Then
                    
                       txt1 = wdDoc.Tables(iTable).Cell(row, iCols1).Range.copy
                       backColor = wdDoc.Tables(iTable).Cell(row, iCols).Shading.BackgroundPatternColor
                       foreColor = wdDoc.Tables(iTable).Cell(row, iCols).Shading.ForegroundPatternColor
                       
                       
                        
                       myCopy.Tables(1).Cell(2, iCols).Range.Italic = False
                       myCopy.Tables(1).Cell(2, iCols).Range.Font.ColorIndex = wdBlack
                       
                       myCopy.Tables(1).Cell(2, iCols).Range.Paste
                       myCopy.Tables(1).Cell(2, iCols).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
                        myCopy.Tables(1).Cell(2, iCols).Range.Font.Size = myCopy.Tables(1).Cell(1, iCols).Range.Font.Size
                       oTable.Cell(2, iCols).Shading.BackgroundPatternColor = backColor ' Paste the Backcolor
                       oTable.Cell(2, iCols).Shading.ForegroundPatternColor = foreColor
                      
                       
                   ElseIf StrComp(txt, s1, vbTextCompare) And txt Like "*Risk Level*" And s1 Like "*Risk Level*" Then
                    
                      txt1 = wdDoc.Tables(iTable).Cell(row, iCols1).Range.copy
                       backColor = wdDoc.Tables(iTable).Cell(row, iCols).Shading.BackgroundPatternColor
                       foreColor = wdDoc.Tables(iTable).Cell(row, iCols).Shading.ForegroundPatternColor
                      
                       myCopy.Tables(1).Cell(2, iCols).Range.Paste
                       myCopy.Tables(1).Cell(2, iCols).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
                       myCopy.Tables(1).Cell(2, iCols).Range.Font.Size = myCopy.Tables(1).Cell(1, iCols).Range.Font.Size
                       myCopy.Tables(1).Cell(2, iCols).Range.Italic = False
                       myCopy.Tables(1).Cell(2, iCols).Range.Font.ColorIndex = wdBlack
                       
                       
                       oTable.Cell(2, iCols).Shading.BackgroundPatternColor = backColor ' Paste the Backcolor
                       oTable.Cell(2, iCols).Shading.ForegroundPatternColor = foreColor
                       
                 
                       
                   Else
                      
                   End If
                    Next iCols1
                    Next iCols
                    
                
                    For iCols = 1 To copyfilecount
                      If myCopy.Tables(1).Cell(2, iCols).Range.Font.ColorIndex = 10 Then
                       
                       myCopy.Tables(1).Cell(2, iCols).Range.Italic = False
                      End If
                    Next iCols
                    
                   For iCols = 1 To iTotalCols
                     rnge2 = wdDoc.Tables(iTable).Cell(row, 2).Range.Text
                   Next iCols
                        
                        rnge3 = ""
                          
                           For position = 0 To Len(rnge2)
                             Select Case Asc(Mid(rnge2, position, 1))
                               Case 48 To 57  ' Ascii codes for numbers from 0-9
                               rnge3 = rnge3 & Mid(rnge2, position, 1)
                               Case 45  ' Ascii code for -
                               rnge3 = rnge3 & Mid(rnge2, position, 1)
                               Case 65 To 90 'Ascii code for  A to Z
                               rnge3 = rnge3 & Mid(rnge2, position, 1)
                               Case 95  'Ascii code for _

                               rnge3 = rnge3 & Mid(rnge2, position, 1)
                               Case 97 To 122 ' Ascii code for a to z
                               rnge3 = rnge3 & Mid(rnge2, position, 1)

                               Case Else
                               rnge3 = rnge3 & " "
                          
                             End Select
                           Next position
                           
                      rnge3 = RTrim(rnge3)
                     
                     myfile = newpath & rnge3 & ".docx"
                                            
                     myCopy.SaveAs2 filename:=myfile
                     myCopy.Close savechanges:=wdSaveChanges
                      
                    'CurrentProgress = i1 / i2
                    'BarWidth = Progress1.Border.width * CurrentProgress
                    'ProgressPercentage = Round(CurrentProgress * 100, 0)
                    'Progress1.Bar.width = BarWidth
                    'Progress1.Text.Caption = i1 & " URS Documents Processed " & vbNewLine & ProgressPercentage & " % Completed "
                    
                    'DoEvents
                    ione = ione + 1
                    i1 = CLng(ione)
                  End With
                 Else
                 End If
                Next row
              End With
              
              'Progress1.Text1.Caption = Tcount & " out of " & count & " Tables Processed  "
              'DoEvents ' Make the changes to impact with the progress bar
              'Tcount = Tcount + 1
           
                            
             Else
             End If
           Next c
        Next r
     End With
   Next iTable
 End With
End If
 
End Sub

