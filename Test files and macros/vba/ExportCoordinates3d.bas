Attribute VB_Name = "ExportCoordinates3d"
Sub catmain()
    
   On Error Resume Next
    
    Dim docPart         As Document
    Dim myPart          As Part
    Dim hybBodies       As HybridBodies
    Dim hybBody         As HybridBody
    Dim hybShapes       As HybridShapes
    Dim hybShape        As HybridShape
 
    Dim arrXYZ(2)
     Dim s               As Long
    Const Separator As String = ";"
    
    Set docPart = CATIA.ActiveDocument
    'If no doc active
    If Err.Number <> 0 Then
        MsgBox "No Active Document", vbCritical
        Exit Sub
    End If
      
    ' Excel variables
    ' http://www.eng-tips.com/viewthread.cfm?qid=308328
    ' Please add reference: Alt+F11 -> Tools -> References... -> Microsoft Excel Object Library
    Dim Excel As Excel.Application
    Dim myWorkBooks As Excel.workbooks
    Dim myWorkBook As Excel.workbook
    Dim myWorkSheet As Excel.worksheet
    Dim myNumberWorkBooks, i As Integer
    Dim myWorkBookFile As String ' numbers od workbooks
    myWorkBookFile = "Calculations.xlsm" ' Plik do otwarcia
    ' end Excel variable
      
    Dim was(0)
    Set userSel = CATIA.ActiveDocument.Selection
    was(0) = "HybridBody"
    userSel.Clear
    aText = userSel.SelectElement2(was, "Select Geometrical Set", True)
 
    Set hybBody = userSel.Item(1).Value
    Set hybShapes = hybBody.HybridShapes

' Podpiecie do Excela
    On Error GoTo ErrHandler
        Set Excel = GetObject(, "EXCEL.Application")
        Set myWorkBooks = Excel.workbooks
ErrHandler:
        If Err.Number <> 0 Then
            MsgBox "Please note you have to run Excel."
        Exit Sub
    End If

    myNumberWorkBooks = myWorkBooks.Count

    ' unikam bladu logiczny i=1 bo moze nie byc woorkbooka, w sumie zbedne
    If myNumberWorkBooks = 0 Then
        MsgBox "Please note you have to open Excel with woorkbooks."
        Exit Sub
    End If

    ' Ladowanie wlasciwego pliku,
    For i = 1 To myNumberWorkBooks
        If (myWorkBooks(i).Name = myWorkBookFile) Then
            Set myWorkBook = myWorkBooks(i)
            Exit For
        ElseIf (i = myNumberWorkBooks) Then
            MsgBox "Please note you have to open Excel template."
            Exit Sub
        End If
    Next i
    
    ' Pierwszy arku laduje
    Excel.Visible = True
    Set myWorkSheet = myWorkBook.Sheets(3)

    ' Search HybridShapes (Point 3D)
    For s = 1 To hybShapes.Count

                Set hybShape = hybShapes.Item(s)
                   
                    'Extract  coord
                    hybShape.GetCoordinates arrXYZ
                 
                    'MsgBox ((hybShapes.Parent.Name) & Separator & (hybShape.Name) & Separator & _
                    '              arrXYZ(0) & Separator & _
                    '              arrXYZ(1) & Separator & _
                    '              arrXYZ(2) & vbLf)
                    
                    'Debug.Print hybShape.Name & " " & _
                    'arrXYZ(0) & _
                    'arrXYZ(1) & _
                    'arrXYZ(2)
                    'Debug.Print sassd + s
                     myWorkSheet.Range("A" & s).Value2 = hybShape.Name
                     myWorkSheet.Range("B" & s).Value2 = arrXYZ(0)
                     myWorkSheet.Range("C" & s).Value2 = arrXYZ(1)
                     myWorkSheet.Range("D" & s).Value2 = arrXYZ(2)
     Next s
End Sub

