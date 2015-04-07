Attribute VB_Name = "SAP_LIST_CHD"
Public Function SAP_LIST_CHILDREN(objParent As Object, iLevel As Integer)
Dim objBook As Workbook
Dim objSheet As Worksheet

Dim lngRow As Long

Dim strId As String
Dim strName As String
Dim strType As String
Dim strText As String

    Set objBook = ActiveSheet.Parent
    Set objSheet = objBook.Worksheets(2)
    objSheet.Select

    If Not objParent Is Nothing Then
        With objParent
            strId = .ID
            strName = .Name
            strType = .Type
            Select Case UCase(strType)
                Case "GUILABEL", "GUITEXTFIELD", "GUICOMBOBOX", "GUISTATUSBASE"
                    strText = .Text
                Case Else
                    strText = vbNullString
            End Select
        End With
    End If
    
    lngRow = (objSheet.Range("A1").CurrentRegion.Rows.Count + 1)

    With objSheet
        .Cells(lngRow, 1).Value = iLevel
        .Cells(lngRow, 2).Value = strId
        .Cells(lngRow, 3).Value = strName
        .Cells(lngRow, 4).Value = strType
        .Cells(lngRow, 5).Value = strText
    End With

Dim objChildren As Object

On Error Resume Next
    Set objChildren = objParent.Children
On Error GoTo 0
    
    If Not objChildren Is Nothing Then
        Dim objChild As Object
        
        For Each objChild In objChildren
            Call SAP_LIST_CHILDREN(objChild, iLevel + 1)
        Next objChild
    End If
End Function
