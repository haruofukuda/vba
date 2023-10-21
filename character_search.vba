'
' search_character executes procedures along the following
'
'1st: create a keyword list retrieved from excel sheet
'2nd: open respective excel book in "items" folder
'3rd: search whether keyword exists per book and note the surrounding value and link if it finds
'
Sub search_character()

    '[constant value]
    '
    'start position of cell which keyword exists
    Const START_POSITION As String = "A2"
    '
    'search range of longitudinal row
    Const SEARCH_RANGE As String = "A:Z"
    '
    'pattern1 character
    Const PATTERN1 As String = "yes"
    '
    'pattern2 character
    Const PATTERN2 As String = "no"
    
    range(START_POSITION).Select
    
    '[1st step]
    '
    'create a keyword collection
    Dim keyList As New Collection

    'get the cell value
    Dim i As Long
    i = 0
    Dim cellValue As String

    With ActiveCell
        cellValue = .Value

        'extract the key
        Do While cellValue <> ""

            'add keyword
            keyList.Add cellValue
            
            'get the cell value
            i = i + 1
            cellValue = .Offset(i, 0).Value
        Loop
    End With


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
                For j = 1 To keyList.Count Step 1
                
                    Dim myObj As range
                    Set myObj = .range(SEARCH_RANGE).Find(keyList(j), LookAt:=xlPart)
                    
                    'if keyword finds
                    If Not (myObj Is Nothing) Then
                    
                        Dim myAddress As String, tooltip As String
                        myAddress = .Name & "!" & myObj.Address
                        tooltip = ""
                        
                        'set adjacent cell value at bottom or right position
                        Dim bottomValue As Variant, rightValue As Variant
                        bottomValue = myObj.Offset(1, 0).Value
                        rightValue = myObj.Offset(0, 1).Value
                        
                        'check whether the adjacent cell contains PATTERN1
                        Dim p1InBottom As Long, p1InRight As Long
                        p1InBottom = InStr(bottomValue, PATTERN1)
                        p1InRight = InStr(rightValue, PATTERN1)
                        
                        Dim matchCond As Boolean
                        matchCond = True
                        
                        Dim k As Long
                        If p1InBottom > 0 Or p1InRight > 0 Then
                            tooltip = createTooltip(myObj, p1InBottom > 0)
                        Else
                            matchCond = False
                        End If
                        
                        If Not matchCond Then
                        
                            'check with same as PATTERN1
                            Dim p2InBottom As Long, p2InRight As Long
                            p2InBottom = InStr(bottomValue, PATTERN2)
                            p2InRight = InStr(rightValue, PATTERN2)
                            
                            If p2InBottom > 0 Or p2InRight > 0 Then
                                tooltip = createTooltip(myObj, p2InBottom > 0)
                            End If
                        End If
                        
                        'note the hyperlink
                        With ThisWorkbook
                            With .ActiveSheet
                                .Hyperlinks.Add Anchor:=.Cells(lateralRow, i), Address:=filePath, SubAddress:=myAddress, ScreenTip:=tooltip, TextToDisplay:="link"
                            End With
                        End With
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


Function createTooltip(object, Optional isBottom As Boolean = False) As String

    'maximum count of offset
    Const OFFSET_COUNT = 5
    
    Dim stringList(OFFSET_COUNT) As String
    For k = 1 To OFFSET_COUNT Step 1
        
        With object
            If isBottom Then
                stringList(k - 1) = .Offset(k, 0).Value
            Else
                stringList(k - 1) = .Offset(0, k).Value
            End If
        End With
    Next k
    
    createTooltip = Join(stringList, "/")
End Function