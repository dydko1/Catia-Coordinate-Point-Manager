Attribute VB_Name = "LeaderCoordinate"
Option Explicit
Sub catmain()

    Dim myview As DrawingView
    Set myview = CATIA.ActiveDocument.Sheets.ActiveSheet.Views.ActiveView
    
    Dim mytext As DrawingText
    Set mytext = myview.Texts.Item(1)
    Dim myleader As DrawingLeader
    
    'Debug.Print mytext.Text
    'Set myleader = mytext.Leaders.Item(1)
    
    Dim oX As Double
    Dim oY As Double
    Dim bShowMove As Boolean
    Dim i As Integer
    bShowMove = True ' pokazuje pozycje lub przesuwa
    
    For i = 1 To myview.Texts.Count
        Set mytext = myview.Texts.Item(i)
        If mytext.Leaders.Count = 1 Then
            Set myleader = mytext.Leaders.Item(1)
            If bShowMove Then
                myleader.GetPoint 1, oX, oY
                Debug.Print ("Point: " & i & Chr(9) & _
                "Text: " & mytext.Text & Chr(9) & Chr(9) & _
                "X: " & oX & Chr(9) & Chr(9) & "Y: " & oY)
            Else
                myleader.ModifyPoint 1, i * 10, i * 10
            End If
        End If
    Next i
End Sub



