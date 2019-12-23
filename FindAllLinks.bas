Attribute VB_Name = "NewMacros"
'
' FindAllLinks Macro and export to Excel file
' Written by Jonny Valai ;)
'

Public i As Integer
Public link_cnt As Integer
Public oExcel
Public oExcelDoc

Sub FindAllLinks()

Set oDoc = ThisDocument

Dim oWord
Dim oLinks
Dim checkLinks
Dim oStoryRanges

Set oExcel = CreateObject("Excel.Application")
Set oExcelDoc = oExcel.Workbooks.Add

Set oLinks = ActiveDocument.Hyperlinks
Set oStoryRanges = ActiveDocument.StoryRanges

checkLinks = MsgBox("This macro will extract and dump all links inside this Word document to a new Excel spreadsheet." & vbCr & vbCr & "Would you like the macro to perform an online check for broken links?", vbYesNo)

oExcel.Visible = True
oExcel.EnableCancelKey = xlInterrupt

oExcel.StatusBar = "Extracting Links... "

With oExcelDoc.Worksheets(1)
    .Range("A1").Value = "Scanning document..."
    .Range("A3").Value = "Status"
    .Range("B3").Value = "Text"
    .Range("C3").Value = "Link"
    .Range("A1:C3").Font.FontStyle = "Bold"
    .Range("A1").Font.Size = 13
End With

i = 4
link_cnt = 0

For Each oStoryRange In oStoryRanges
    WriteLinks oStoryRange.Hyperlinks, checkLinks
Next oStoryRange

'WriteLinks oLinks, checkLinks

'Find links inside shapes, textboxes etc
GetShapeLinks checkLinks

With oExcelDoc.Worksheets(1)
    .Range("A1").Value = "Found " & link_cnt & " hyperlinks in file '" & ActiveDocument.Name & "'"
End With

oExcel.StatusBar = "Finished scanning."

End Sub


Sub WriteLinks(oLinks, checkLinks)

Dim oHTTP

On Error Resume Next

For Each oLink In oLinks

    If oLink.Address <> "" Then
    
        If checkLinks = vbYes Then
            oExcel.StatusBar = "Testing Link: " & oLink.Address
            
            Set oHTTP = CreateObject("MSXML2.XMLHTTP")
            oHTTP.Open "HEAD", oLink.Address, False
            
            oHTTP.Send
        
            If checkLinks = vbYes Then
                If oHTTP.StatusText = "" Then
                    oLink.Range.HighlightColorIndex = wdColorOrange
                Else
                    oLink.Range.HighlightColorIndex = wdColorGreen
                End If
                
                If oHTTP.StatusText = "Not Found" Then
                    oLink.Range.HighlightColorIndex = wdColorRed
                End If
               
            End If
        
        End If
        
        With oExcelDoc.Worksheets(1)
            .Range("B" & i).Value = oLink.TextToDisplay
            .Hyperlinks.Add Anchor:=oExcelDoc.Worksheets(1).Range("C" & i), Address:=oLink.Address, TextToDisplay:=oLink.Address
            
            If checkLinks = vbYes Then
                If oHTTP.StatusText = "" Then
                    .Range("A" & i & ":C" & i).Font.Color = wdColorOrange
                    .Range("A" & i).Value = "?"
                Else
                    .Range("A" & i & ":C" & i).Font.Color = wdColorGreen
                    .Range("A" & i).Value = oHTTP.StatusText
                End If
                
                If oHTTP.StatusText = "Not Found" Then
                    .Range("A" & i & ":C" & i).Font.Color = wdColorRed
                End If
               
            End If
            
        End With
    
        oExcel.StatusBar = False
    
        i = i + 1
        
        link_cnt = link_cnt + 1
        
    End If
    
Next

End Sub

Sub GetShapeLinks(checkLinks)

    Dim oShapes
    Set oShapes = ActiveDocument.Shapes

    For Each oShape In oShapes
    
        If oShape.Type = msoGroup Then
        
            On Error Resume Next
            Set oShapeRange = oShape.Ungroup
            
            For Each oGroupMember In oShapeRange
                
                If oGroupMember.Type = msoGroup Then
                    Set oShapeRange2 = oGroupMember.Ungroup
                End If
            
            Next oGroupMember
            
        End If
    
    Next oShape
    
    Set oShapes = ActiveDocument.Shapes
    For Each oShape In oShapes
        
       oShape.Select
      
       WriteLinks Selection.Hyperlinks, checkLinks
        
    Next oShape

End Sub


