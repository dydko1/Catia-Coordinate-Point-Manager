Attribute VB_Name = "GetVectors"
Option Explicit
Sub catmain()
    ' http://ww3.cad.de/foren/ubb/Forum137/HTML/002783.shtml
    ' http://www.eng-tips.com/viewthread.cfm?qid=308328
    ' Catia varaibles
    Dim oActDoc As Document
    Set oActDoc = CATIA.ActiveDocument
    ' Dim oSheet As DrawingSheet
    ' Set oSheet = oActDoc.Sheets.Item(1)
    Dim oView As DrawingView
    ' Dim iView As Integer
    ' iView = 4 ' aktualny widok = 2 + iView
    ' Set oView = oSheet.Views.Item(2 + iView)
    Set oView = oActDoc.Sheets.ActiveSheet.Views.ActiveView ' bez liczenia
    ' end catia variable
    
    ' Excel variables
    ' Please add reference: Alt+F11 -> Tools -> References... -> Microsoft Excel Object Library
    Dim Excel As Excel.Application
    Dim myWorkBooks As Excel.workbooks
    Dim myWorkBook As Excel.workbook
    Dim myWorkSheet As Excel.worksheet
    Dim myNumberWorkBooks, i As Integer
    Dim myWorkBookFile As String ' numbers od workbooks
    myWorkBookFile = "Calculations.xlsm" ' Plik do otwarcia
    ' end Excel variable

    Dim oGenBeh As DrawingViewGenerativeBehavior
    Set oGenBeh = oView.GenerativeBehavior

    Dim X1x As Double ' wersor U
    Dim X1y As Double
    Dim X1z As Double
    Dim Y1x As Double ' wersor V
    Dim Y1y As Double
    Dim Y1z As Double
    oGenBeh.GetProjectionPlane X1x, X1y, X1z, Y1x, Y1y, Y1z
    ' Debug.Print "U= " & X1x, X1y, X1z
    ' Debug.Print "V= " & Y1x, Y1y, Y1z
    
    Dim Z1x As Double
    Dim Z1y As Double
    Dim Z1z As Double
    oGenBeh.GetProjectionPlaneNormal Z1x, Z1y, Z1z
    'Debug.Print "N= " & Z1x, Z1y, Z1z
    
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
    Set myWorkSheet = myWorkBook.Sheets(1)
       
    ' nazwa aktualnego view
    myWorkSheet.Range("B11").Value2 = oView.Name
    ' wersory
    myWorkSheet.Range("C13").Value2 = X1x
    myWorkSheet.Range("D13").Value2 = X1y
    myWorkSheet.Range("E13").Value2 = X1z
    myWorkSheet.Range("C14").Value2 = Y1x
    myWorkSheet.Range("D14").Value2 = Y1y
    myWorkSheet.Range("E14").Value2 = Y1z
    myWorkSheet.Range("C15").Value2 = Z1x
    myWorkSheet.Range("D15").Value2 = Z1y
    myWorkSheet.Range("E15").Value2 = Z1z
End Sub

