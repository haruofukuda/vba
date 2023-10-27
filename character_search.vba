'[constant value]
'
'start position of cell which keyword exists
Private Const START_POSITION As String = "A2"
Private Const START_POSITION2 As String = "A7"
'
'search range of longitudinal row
Private Const SEARCH_RANGE As String = "A:Z"
'
'maximum count of offset
Private Const OFFSET_COUNT As Integer = 5

'
' search_character Sub statement executes procedures along the following
'
'1st : create a keyword list retrieved from excel sheet
'2nd : open respective excel book in "items" folder
'3rd : search whether keyword exists per the book
'Goal: note the surrounding value and link if it finds
'
Sub search_character()

    range(START_POSITION).Select
    
    '[1st step]
    '
    'create a keyword collection
    Dim keyList
    Set keyList = collectWords()

    range(START_POSITION2).Select

    Dim selectList
    Set selectList = collectWords()

    '[2nd step]
    '
    'specify a folder which opens at first
    Dim folderName As String
    folderName = ThisWorkbook.Path & "\items"
    
    'get a file name
    Dim fileName As String
    fileName = Dir(folderName & "\*")
    
    i = 2
    Do While fileName <> ""
    
        'open the file
        Dim filePath As String
        filePath = folderName & "\" & fileName
        Workbooks.Open fileName:=filePath
        
        'note the file name
        Dim lateralRow As Long
        lateralRow = 1
        ThisWorkbook.ActiveSheet.Cells(lateralRow, i) = ActiveWorkbook.Name
        lateralRow = lateralRow + 1
        
        
        '[3rd step]
        '
        'search keywords
        With ActiveWorkbook
            With .ActiveSheet
                
                Dim j As Long
                For j = 1 To keyList.count Step 1
                
                    Dim myObj As range
                    Set myObj = .range(SEARCH_RANGE).Find(keyList(j), LookAt:=xlPart)
                    
                    'if keyword finds
                    If Not (myObj Is Nothing) Then
                    
                        Dim myAddress As String, seriesList() As String, linkText As String, pattern1 As String, pattern2 As String
                        myAddress = .Name & "!" & myObj.Address
                        linkText = "link"
                        pattern1 = selectList(1)
                        pattern2 = selectList(2)
                        
                        'set adjacent cell value at bottom or right position
                        Dim bottomValue As Variant, rightValue As Variant
                        bottomValue = myObj.Offset(1, 0).Value
                        rightValue = myObj.Offset(0, 1).Value
                        
                        'check whether the adjacent cell contains PATTERN1
                        Dim p1InBottom As Long, p1InRight As Long
                        p1InBottom = InStr(bottomValue, pattern1)
                        p1InRight = InStr(rightValue, pattern1)
                        
                        Dim matchCond As Boolean
                        matchCond = True
                        
                        Dim k As Long
                        If p1InBottom > 0 Or p1InRight > 0 Then
                            seriesList = adjacentList(myObj, p1InBottom > 0)
                        Else
                            matchCond = False
                        End If
                        
                        If Not matchCond Then
                        
                            'check with same as PATTERN1
                            Dim p2InBottom As Long, p2InRight As Long
                            p2InBottom = InStr(bottomValue, pattern2)
                            p2InRight = InStr(rightValue, pattern2)
                            
                            If p2InBottom > 0 Or p2InRight > 0 Then
                                seriesList = adjacentList(myObj, p2InBottom > 0)
                            End If
                        End If
                        
                        If (Not seriesList) <> -1 Then
                            Dim maxLen As Long
                            maxLen = 0
                            For k = LBound(seriesList) To UBound(seriesList) Step 1
                                If Len(seriesList(k)) > maxLen Then
                                    linkText = seriesList(k)
                                    maxLen = Len(seriesList(k))
                                End If
                            Next k
                        
                            '[Goal]
                            '
                            'note the hyperlink
                            With ThisWorkbook
                                With .ActiveSheet
                                    .Hyperlinks.Add Anchor:=.Cells(lateralRow, i), Address:=filePath, SubAddress:=myAddress, ScreenTip:=Join(seriesList, "/"), TextToDisplay:=linkText
                                    
                                    'initialize the adjacent text list
                                    Erase seriesList
                                End With
                            End With
                        End If
                    End If
                    
                    lateralRow = lateralRow + 1
                Next j
            End With
            
            'close the file
            .Close
        End With
        
        fileName = Dir()
        i = i + 1
    Loop
End Sub

Function collectWords() As Collection

    Dim arr As Collection
    Set arr = New Collection
    
    'get the cell value
    Dim i As Long
    i = 0
    Dim cellValue As String

    With ActiveCell
        cellValue = .Value

        'extract the key
        Do While cellValue <> ""

            'add keyword
            arr.Add cellValue
            
            'get the cell value
            i = i + 1
            cellValue = .Offset(i, 0).Value
        Loop
    End With
    
    Set collectWords = arr
End Function

Function adjacentList(object, Optional isBottom As Boolean = False) As String()
    
    Dim stringList(OFFSET_COUNT) As String
    For k = 1 To OFFSET_COUNT - 1 Step 1
        
        With object
            If isBottom Then
                stringList(k - 1) = .Offset(k, 0).Value
            Else
                stringList(k - 1) = .Offset(0, k).Value
            End If
        End With
    Next k
    
    adjacentList = stringList
End Function