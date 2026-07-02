Attribute VB_Name = "TubingConsolidator"
Option Explicit

' ====================================================================
' AUTODESK INVENTOR VBA TUBING CONSOLIDATION PROGRAM
' Purpose: Consolidates tubing parts from an assembly into a drawing sheet
' Author: Generated for Winholt Equipment
' ====================================================================

' Global objects
Dim invApp As Inventor.Application
Dim invDoc As Inventor.Document
Dim invDrawing As Inventor.DrawingDocument
Dim invSheet As Inventor.Sheet
Dim tubesByIBM As Object ' Dictionary: IBM -> Dictionary(SubAsmName -> Collection of TubeData)
Dim processedDocs As Object ' Collection to track processed documents

' Configuration variables (user-adjustable)
Dim VAR_BASE_Y_POS As Double
Dim VAR_TUBE_SPACING As Double
Dim VAR_BASE_X_POS As Double
Dim VAR_CROSS_SECTION_OFFSET_X As Double
Dim VAR_PROJECTION_OFFSET_X As Double
Dim VAR_PARTS_LIST_OFFSET_X As Double

' ====================================================================
' MAIN ENTRY POINT
' ====================================================================
Sub Main()
    On Error GoTo ErrorHandler
    
    ' Initialize configuration variables
    InitializeConfiguration
    
    ' Initialize Inventor application
    Set invApp = ThisDrawing.Application
    Set invDoc = invApp.ActiveDocument
    
    ' Validate document type
    If invDoc.DocumentType <> kDrawingDocumentObject Then
        MsgBox "Please open a Drawing document first.", vbCritical
        Exit Sub
    End If
    
    Set invDrawing = invDoc
    
    ' Initialize collections
    Set tubesByIBM = CreateObject("Scripting.Dictionary")
    Set processedDocs = CreateObject("Scripting.Dictionary")
    
    ' Step 1: Get the main assembly from the first sheet
    Dim mainAssemblyDoc As Inventor.Document
    Set mainAssemblyDoc = GetAssemblyFromFirstSheet(invSheet)
    
    If mainAssemblyDoc Is Nothing Then
        MsgBox "Could not find assembly on first sheet.", vbCritical
        Exit Sub
    End If
    
    ' Step 2: Collect tubing parts organized by IBM and sub-assembly
    CollectTubingParts mainAssemblyDoc, tubesByIBM, processedDocs
    
    ' Validate that tubes were found
    If tubesByIBM.Count = 0 Then
        MsgBox "No tubing parts found (material codes 351-369).", vbExclamation
        Exit Sub
    End If
    
    ' Step 3: Create drawing sheets for each IBM's tubes
    CreateTubingSheetsByIBM invDrawing, tubesByIBM
    
    MsgBox "Tubing consolidation complete! Created " & GetTotalTubingSheets(invDrawing) & " tubing sheet(s).", vbInformation
    
    Exit Sub
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

' ====================================================================
' INITIALIZE CONFIGURATION VARIABLES
' ====================================================================
Sub InitializeConfiguration()
    ' These variables control placement and can be adjusted for manual fine-tuning
    VAR_BASE_Y_POS = 9.0 ' Distance from top of sheet (inches)
    VAR_TUBE_SPACING = 2.2 ' Vertical distance between each tube (inches)
    VAR_BASE_X_POS = 0.75 ' Distance from left edge (inches)
    VAR_CROSS_SECTION_OFFSET_X = 0 ' Additional X offset for cross-section view
    VAR_PROJECTION_OFFSET_X = 1.2 ' Horizontal offset from cross-section to projection
    VAR_PARTS_LIST_OFFSET_X = 1.8 ' Horizontal offset from projection to parts list
End Sub

' ====================================================================
' HELPER: GET ASSEMBLY FROM FIRST SHEET
' ====================================================================
Function GetAssemblyFromFirstSheet(ByRef outSheet As Inventor.Sheet) As Inventor.Document
    On Error GoTo ErrorHandler
    
    Dim sheet As Inventor.Sheet
    Dim view As Inventor.DrawingView
    Dim i As Long
    
    ' Get the first sheet
    Set sheet = invDrawing.Sheets.Item(1)
    Set outSheet = sheet
    
    ' Loop through views on the first sheet to find an assembly view
    For i = 1 To sheet.DrawingViews.Count
        Set view = sheet.DrawingViews.Item(i)
        
        ' Check if view references an assembly document
        On Error Resume Next
        If view.ReferencedDocumentDescriptor.ReferencedDocument.DocumentType = kAssemblyDocumentObject Then
            Set GetAssemblyFromFirstSheet = view.ReferencedDocumentDescriptor.ReferencedDocument
            Exit Function
        End If
        On Error GoTo ErrorHandler
    Next i
    
    Exit Function
ErrorHandler:
    Set GetAssemblyFromFirstSheet = Nothing
End Function

' ====================================================================
' STEP 1: COLLECT TUBING PARTS ORGANIZED BY IBM
' ====================================================================
Sub CollectTubingParts(asmDoc As Inventor.Document, ibmDict As Object, procDocs As Object)
    On Error GoTo ErrorHandler
    
    Dim asmDef As Inventor.AssemblyComponentDefinition
    Dim occ As Inventor.ComponentOccurrence
    Dim i As Long
    Dim materialCode As String
    Dim materialNum As Integer
    Dim tubeData As TubeData
    Dim ibmNum As Long
    Dim addendumNum As Long
    Dim fileName As String
    Dim subAsmDict As Object
    Dim tubeColl As Collection
    
    ' Get the assembly definition
    If asmDoc.DocumentType <> kAssemblyDocumentObject Then
        Exit Sub
    End If
    
    Set asmDef = asmDoc.ComponentDefinition
    
    ' Loop through all occurrences in the assembly
    For i = 1 To asmDef.Occurrences.Count
        Set occ = asmDef.Occurrences.Item(i)
        
        ' Process recursively if this is a sub-assembly
        If occ.Definition.Document.DocumentType = kAssemblyDocumentObject Then
            If Not procDocs.Exists(occ.Definition.Document.FullFileName) Then
                procDocs.Add occ.Definition.Document.FullFileName, True
                CollectTubingParts occ.Definition.Document, ibmDict, procDocs
            End If
        Else
            ' This is a part - check if it's a valid tubing part
            fileName = GetFileNameWithoutExtension(occ.Definition.Document.FullFileName)
            
            ' Get IBM and Addendum numbers
            ibmNum = GetIBMNum(fileName)
            addendumNum = GetAddendumNum(fileName)
            
            ' Validate this is a valid part (has both IBM and Addendum)
            If ibmNum > 0 And addendumNum >= 0 Then
                ' Check material code
                On Error Resume Next
                materialCode = GetIProperty(occ.Definition.Document, "Physical", "Material")
                On Error GoTo ErrorHandler
                
                ' Check if material code is tubing (351-369)
                If Len(materialCode) >= 3 Then
                    materialNum = CInt(Left(materialCode, 3))
                    If materialNum >= 351 And materialNum <= 369 Then
                        ' Create tube data structure
                        tubeData.PartName = occ.Name
                        tubeData.PartNumber = fileName
                        tubeData.Description = GetIProperty(occ.Definition.Document, "Summary", "Title")
                        tubeData.Material = materialCode
                        tubeData.Document = occ.Definition.Document
                        tubeData.ComponentOcc = occ
                        tubeData.IBMNum = ibmNum
                        tubeData.AddendumNum = addendumNum
                        
                        ' Create IBM dictionary if it doesn't exist
                        If Not ibmDict.Exists(CStr(ibmNum)) Then
                            ibmDict.Add CStr(ibmNum), CreateObject("Scripting.Dictionary")
                        End If
                        
                        Set subAsmDict = ibmDict.Item(CStr(ibmNum))
                        
                        ' Create sub-assembly collection if it doesn't exist
                        Dim subAsmKey As String
                        subAsmKey = "Main"
                        
                        If Not subAsmDict.Exists(subAsmKey) Then
                            Set tubeColl = New Collection
                            subAsmDict.Add subAsmKey, tubeColl
                        End If
                        
                        ' Add to appropriate collection
                        subAsmDict.Item(subAsmKey).Add tubeData
                    End If
                End If
            End If
        End If
    Next i
    
    Exit Sub
ErrorHandler:
    ' Silently continue on error
End Sub

' ====================================================================
' HELPER: GET FILE NAME WITHOUT EXTENSION
' ====================================================================
Function GetFileNameWithoutExtension(fullPath As String) As String
    Dim fileName As String
    Dim lastSlash As Long
    Dim lastDot As Long
    
    ' Get file name from path
    lastSlash = InStrRev(fullPath, "\")
    If lastSlash = 0 Then lastSlash = InStrRev(fullPath, "/")
    fileName = Mid(fullPath, lastSlash + 1)
    
    ' Remove extension
    lastDot = InStrRev(fileName, ".")
    If lastDot > 0 Then
        fileName = Left(fileName, lastDot - 1)
    End If
    
    GetFileNameWithoutExtension = fileName
End Function

' ====================================================================
' HELPER: GET IBM NUMBER
' ====================================================================
Function GetIBMNum(fileName As String) As Long
    Dim parts() As String
    Dim splitDash() As String
    
    ' Split by space first
    parts = Split(fileName, " ")
    If UBound(parts) < 0 Then
        GetIBMNum = -1
        Exit Function
    End If
    
    ' Split first part by hyphen
    splitDash = Split(parts(0), "-")
    If UBound(splitDash) <> 1 Then
        GetIBMNum = -1
        Exit Function
    End If
    
    ' Validate IBM number
    If Len(splitDash(0)) < 6 Then
        GetIBMNum = -1
        Exit Function
    End If
    
    If Not IsNumeric(splitDash(0)) Then
        GetIBMNum = -1
        Exit Function
    End If
    
    GetIBMNum = CLng(splitDash(0))
End Function

' ====================================================================
' HELPER: GET ADDENDUM NUMBER
' ====================================================================
Function GetAddendumNum(fileName As String) As Long
    Dim parts() As String
    Dim splitDash() As String
    
    ' Split by space first
    parts = Split(fileName, " ")
    If UBound(parts) < 0 Then
        GetAddendumNum = -1
        Exit Function
    End If
    
    ' Split first part by hyphen
    splitDash = Split(parts(0), "-")
    If UBound(splitDash) <> 1 Then
        GetAddendumNum = -1
        Exit Function
    End If
    
    ' Validate IBM part
    If Len(splitDash(0)) < 6 Or Not IsNumeric(splitDash(0)) Then
        GetAddendumNum = -1
        Exit Function
    End If
    
    ' Try to parse addendum number
    If IsNumeric(splitDash(1)) Then
        GetAddendumNum = CLng(splitDash(1))
    Else
        GetAddendumNum = -1
    End If
End Function

' ====================================================================
' HELPER: CHECK IF STRING IS NUMERIC
' ====================================================================
Function IsNumeric(str As String) As Boolean
    On Error Resume Next
    IsNumeric = Not IsNull(CLng(str))
    On Error GoTo 0
End Function

' ====================================================================
' HELPER: GET IPROPERTY VALUE
' ====================================================================
Function GetIProperty(doc As Inventor.Document, section As String, property As String) As String
    On Error Resume Next
    GetIProperty = doc.PropertySets.Item(section).Item(property).Value
    If Err.Number <> 0 Then GetIProperty = ""
    On Error GoTo 0
End Function

' ====================================================================
' STEP 2 & 3: CREATE TUBING SHEETS FOR EACH IBM
' ====================================================================
Sub CreateTubingSheetsByIBM(drawDoc As Inventor.DrawingDocument, ibmDict As Object)
    Dim ibmKey As Variant
    Dim subAsmDict As Object
    Dim subAsmKey As Variant
    Dim tubeColl As Collection
    Dim sheetCount As Integer
    Dim tubesPerSheet As Integer
    Dim currentSheet As Integer
    Dim i As Long
    Dim startIdx As Long
    Dim endIdx As Long
    Dim scale As Double
    Dim ibmNum As Long
    
    tubesPerSheet = 6
    
    ' Iterate through each IBM
    For Each ibmKey In ibmDict.Keys
        ibmNum = CLng(ibmKey)
        Set subAsmDict = ibmDict.Item(ibmKey)
        
        ' Iterate through each sub-assembly within this IBM
        For Each subAsmKey In subAsmDict.Keys
            Set tubeColl = subAsmDict.Item(subAsmKey)
            
            ' Sort tubes by part number
            SortTubesByPartNumber tubeColl
            
            ' Calculate number of sheets needed for this sub-assembly
            sheetCount = Int((tubeColl.Count - 1) / tubesPerSheet) + 1
            
            ' Create sheets for this sub-assembly
            For currentSheet = 1 To sheetCount
                startIdx = (currentSheet - 1) * tubesPerSheet + 1
                endIdx = Application.Min(startIdx + tubesPerSheet - 1, tubeColl.Count)
                
                ' Create new sheet
                Set invSheet = CreateNewTubingSheet(drawDoc, ibmNum, CStr(subAsmKey), currentSheet, sheetCount)
                
                ' Determine scale based on tube count
                scale = DetermineScale(endIdx - startIdx + 1)
                
                ' Place tubes on sheet
                PlaceTubesOnSheet drawDoc, invSheet, tubeColl, startIdx, endIdx, scale
            Next currentSheet
        Next subAsmKey
    Next ibmKey
End Sub

' ====================================================================
' HELPER: SORT TUBES BY PART NUMBER
' ====================================================================
Sub SortTubesByPartNumber(tubeColl As Collection)
    Dim i As Long, j As Long
    Dim tubeA As TubeData, tubeB As TubeData
    Dim aAddendum As Long, bAddendum As Long
    
    ' Bubble sort by addendum number
    For i = 1 To tubeColl.Count - 1
        For j = i + 1 To tubeColl.Count
            tubeA = tubeColl.Item(i)
            tubeB = tubeColl.Item(j)
            
            aAddendum = tubeA.AddendumNum
            bAddendum = tubeB.AddendumNum
            
            If aAddendum > bAddendum Then
                ' Swap
                tubeColl.Remove j
                tubeColl.Add tubeA, , j
            End If
        Next j
    Next i
End Sub

' ====================================================================
' HELPER: CREATE NEW TUBING SHEET
' ====================================================================
Function CreateNewTubingSheet(drawDoc As Inventor.DrawingDocument, ibmNum As Long, _
                              subAsmName As String, sheetNum As Integer, totalSheets As Integer) As Inventor.Sheet
    Dim newSheet As Inventor.Sheet
    Dim sheetName As String
    Dim tg As Inventor.TransientGeometry
    
    Set tg = invApp.TransientGeometry
    
    ' Create sheet name
    If totalSheets > 1 Then
        sheetName = "LEGS & SPREADERS - " & ibmNum & " (" & sheetNum & "/" & totalSheets & ")"
    Else
        sheetName = "LEGS & SPREADERS - " & ibmNum
    End If
    
    ' Create new sheet
    Set newSheet = drawDoc.Sheets.Add
    newSheet.Name = sheetName
    
    ' Try to apply WINHOLT title block if available
    On Error Resume Next
    Dim titleBlockDef As Inventor.TitleBlockDefinition
    Set titleBlockDef = drawDoc.TitleBlockDefinitions.Item("WINHOLT")
    
    If Not newSheet.TitleBlock Is Nothing Then
        newSheet.TitleBlock.Delete
    End If
    
    Dim sPromptStrings(0 To 2) As String
    sPromptStrings(0) = ""
    sPromptStrings(1) = "LEGS & SPREADERS"
    sPromptStrings(2) = ""
    
    newSheet.AddTitleBlock titleBlockDef, , sPromptStrings
    On Error GoTo 0
    
    Set CreateNewTubingSheet = newSheet
End Function

' ====================================================================
' HELPER: DETERMINE SCALE BASED ON TUBE COUNT
' ====================================================================
Function DetermineScale(tubeCount As Integer) As Double
    Select Case tubeCount
        Case 2, 3
            DetermineScale = 0.25
        Case 4, 5, 6
            DetermineScale = 0.1875
        Case Else
            DetermineScale = 0.1875
    End Select
End Function

' ====================================================================
' STEP 4: PLACE TUBES ON SHEET
' ====================================================================
Sub PlaceTubesOnSheet(drawDoc As Inventor.DrawingDocument, sheet As Inventor.Sheet, _
                      tubeColl As Collection, startIdx As Long, endIdx As Long, scale As Double)
    Dim i As Long
    Dim tubeData As TubeData
    Dim yPos As Double
    Dim xPos As Double
    Dim view As Inventor.DrawingView
    Dim projView As Inventor.DrawingView
    Dim tg As Inventor.TransientGeometry
    Dim crossSectionX As Double
    
    Set tg = invApp.TransientGeometry
    
    crossSectionX = VAR_BASE_X_POS + VAR_CROSS_SECTION_OFFSET_X
    
    ' Place each tube
    For i = startIdx To endIdx
        tubeData = tubeColl.Item(i)
        yPos = VAR_BASE_Y_POS - ((i - startIdx) * VAR_TUBE_SPACING)
        
        ' Create cross-section view
        On Error Resume Next
        Set view = PlaceCrossSectionView(drawDoc, sheet, tubeData, crossSectionX, yPos, scale, tg)
        On Error GoTo 0
        
        If Not view Is Nothing Then
            ' Create projection view with length
            On Error Resume Next
            Set projView = PlaceLengthProjectionView(drawDoc, sheet, view, tubeData, scale, tg)
            On Error GoTo 0
            
            If Not projView Is Nothing Then
                ' Add length dimension to projection view
                AddLengthDimension sheet, projView, tubeData, tg
                
                ' Add parts list box
                AddPartsList sheet, tubeData, projView, tg
                
                ' Check and add break line if necessary
                If NeedsBreakLine(tubeData, scale, sheet) Then
                    AddBreakLine sheet, projView
                End If
            End If
        End If
    Next i
End Sub

' ====================================================================
' HELPER: PLACE CROSS-SECTION VIEW
' ====================================================================
Function PlaceCrossSectionView(drawDoc As Inventor.DrawingDocument, sheet As Inventor.Sheet, _
                               tubeData As TubeData, xPos As Double, yPos As Double, scale As Double, _
                               tg As Inventor.TransientGeometry) As Inventor.DrawingView
    Dim view As Inventor.DrawingView
    Dim placement As Inventor.Point2d
    
    Set placement = tg.CreatePoint2d(xPos, yPos)
    
    ' Create base view from tube part
    ' Right view orientation shows the cross-section (smaller surface area)
    Set view = sheet.DrawingViews.AddBaseView(tubeData.Document, placement, scale, kRightViewOrientation)
    
    Set PlaceCrossSectionView = view
End Function

' ====================================================================
' HELPER: PLACE LENGTH PROJECTION VIEW
' ====================================================================
Function PlaceLengthProjectionView(drawDoc As Inventor.DrawingDocument, sheet As Inventor.Sheet, _
                                   baseView As Inventor.DrawingView, tubeData As TubeData, scale As Double, _
                                   tg As Inventor.TransientGeometry) As Inventor.DrawingView
    Dim projView As Inventor.DrawingView
    Dim projPlacement As Inventor.Point2d
    Dim projXPos As Double
    
    ' Position projection to the right of cross-section
    projXPos = baseView.Position.X + VAR_PROJECTION_OFFSET_X
    Set projPlacement = tg.CreatePoint2d(projXPos, baseView.Position.Y)
    
    ' Create projection view (top view to show length)
    Set projView = sheet.DrawingViews.AddProjectedView(baseView, projPlacement, kDefaultViewLabel)
    
    Set PlaceLengthProjectionView = projView
End Function

' ====================================================================
' HELPER: ADD LENGTH DIMENSION
' ====================================================================
Sub AddLengthDimension(sheet As Inventor.Sheet, projView As Inventor.DrawingView, _
                       tubeData As TubeData, tg As Inventor.TransientGeometry)
    On Error GoTo ErrorHandler
    
    Dim tubeLength As Double
    Dim edges As Object
    Dim edge As Inventor.Edge
    Dim edgeLength As Double
    Dim maxLength As Double
    Dim i As Long
    Dim dimPoint1 As Inventor.Point2d
    Dim dimPoint2 As Inventor.Point2d
    Dim dim As Inventor.DrawingDimension
    
    maxLength = 0
    
    ' Get the part definition
    Dim partDef As Inventor.PartComponentDefinition
    Set partDef = tubeData.Document.ComponentDefinition
    
    ' Find the longest edge (likely the tube length)
    ' Iterate through edges to find longest linear edge
    For Each edge In partDef.SurfaceBodies(1).Edges
        On Error Resume Next
        If edge.GeometryType = kLinearCurveObject Then
            edgeLength = edge.Curve.GetLength()
            If edgeLength > maxLength Then
                maxLength = edgeLength
            End If
        End If
        On Error GoTo ErrorHandler
    Next edge
    
    ' If we found a length, create dimension points
    If maxLength > 0 Then
        ' Create dimension at the end of the projection view
        Set dimPoint1 = tg.CreatePoint2d(projView.Position.X + 0.2, projView.Position.Y - 0.3)
        Set dimPoint2 = tg.CreatePoint2d(projView.Position.X + 0.2, projView.Position.Y + 0.3)
        
        ' Add text annotation with length
        ' Note: Actual dimension placement would need curve references from the view
        ' This is a placeholder - can be enhanced with proper dimension handling
    End If
    
    Exit Sub
ErrorHandler:
    ' Silently fail - dimension can be added manually
End Sub

' ====================================================================
' HELPER: ADD PARTS LIST
' ====================================================================
Sub AddPartsList(sheet As Inventor.Sheet, tubeData As TubeData, refView As Inventor.DrawingView, _
                 tg As Inventor.TransientGeometry)
    On Error Resume Next
    
    Dim placement As Inventor.Point2d
    Dim xPos As Double
    Dim yPos As Double
    Dim textObj As Inventor.TextBox
    Dim partsListText As String
    
    ' Position parts list to right of projection view
    xPos = refView.Position.X + VAR_PARTS_LIST_OFFSET_X
    yPos = refView.Position.Y
    
    Set placement = tg.CreatePoint2d(xPos, yPos)
    
    ' Create text box with parts list information
    ' Format: Three columns - PART NUMBER, DESCRIPTION, MATERIAL
    partsListText = "PART NUMBER" & vbCrLf & tubeData.PartNumber & vbCrLf & vbCrLf & _
                    "DESCRIPTION" & vbCrLf & tubeData.Description & vbCrLf & vbCrLf & _
                    "MATERIAL" & vbCrLf & tubeData.Material
    
    ' Add text box to sheet (approximate size for parts list)
    Set textObj = sheet.Sketches.Add.SketchTexts.Add(partsListText, placement)
    textObj.Alignment = kLeftAlignment
    textObj.FontSize = 8
    
    On Error GoTo 0
End Sub

' ====================================================================
' HELPER: CHECK IF BREAK LINE NEEDED
' ====================================================================
Function NeedsBreakLine(tubeData As TubeData, scale As Double, sheet As Inventor.Sheet) As Boolean
    On Error GoTo ErrorHandler
    
    Dim tubeLength As Double
    Dim maxLength As Double
    Dim edge As Inventor.Edge
    Dim edgeLength As Double
    Dim sheetHeight As Double
    Dim maxAllowedLength As Double
    Dim partDef As Inventor.PartComponentDefinition
    
    ' Standard sheet height is 11" (A size)
    sheetHeight = 11
    ' 1/4 of sheet length at current scale
    maxAllowedLength = (sheetHeight / 4) / scale
    
    ' Get the part definition
    Set partDef = tubeData.Document.ComponentDefinition
    
    maxLength = 0
    
    ' Find the longest edge (tube length)
    For Each edge In partDef.SurfaceBodies(1).Edges
        On Error Resume Next
        If edge.GeometryType = kLinearCurveObject Then
            edgeLength = edge.Curve.GetLength()
            If edgeLength > maxLength Then
                maxLength = edgeLength
            End If
        End If
        On Error GoTo ErrorHandler
    Next edge
    
    ' Return True if tube length exceeds 1/4 of sheet
    If maxLength > maxAllowedLength And maxLength > 0 Then
        NeedsBreakLine = True
    Else
        NeedsBreakLine = False
    End If
    
    Exit Function
ErrorHandler:
    NeedsBreakLine = False
End Function

' ====================================================================
' HELPER: ADD BREAK LINE
' ====================================================================
Sub AddBreakLine(sheet As Inventor.Sheet, projView As Inventor.DrawingView)
    On Error Resume Next
    
    Dim tg As Inventor.TransientGeometry
    Set tg = invApp.TransientGeometry
    
    ' Get the center of the projection view
    Dim centerX As Double
    Dim centerY As Double
    Dim topY As Double
    Dim bottomY As Double
    Dim breakLineStartPt As Inventor.Point2d
    Dim breakLineEndPt As Inventor.Point2d
    
    centerX = projView.Position.X
    centerY = projView.Position.Y
    
    ' Estimate top and bottom of view based on view size
    ' This is approximate - actual break line positioning may need refinement
    topY = centerY + 0.5
    bottomY = centerY - 0.5
    
    ' Create break line points
    Set breakLineStartPt = tg.CreatePoint2d(centerX - 0.1, centerY + 0.2)
    Set breakLineEndPt = tg.CreatePoint2d(centerX + 0.1, centerY - 0.2)
    
    ' Note: Actual break line implementation would use view's break line features
    ' This is a placeholder for the break line logic
    ' In Inventor, break lines are typically added through view properties
    
    On Error GoTo 0
End Sub

' ====================================================================
' HELPER: COUNT TOTAL TUBING SHEETS
' ====================================================================
Function GetTotalTubingSheets(drawDoc As Inventor.DrawingDocument) As Integer
    Dim count As Integer
    Dim i As Long
    
    count = 0
    For i = 1 To drawDoc.Sheets.Count
        If InStr(drawDoc.Sheets.Item(i).Name, "LEGS & SPREADERS") > 0 Then
            count = count + 1
        End If
    Next i
    
    GetTotalTubingSheets = count
End Function

' ====================================================================
' DATA STRUCTURE: TUBE DATA
' ====================================================================
Type TubeData
    PartName As String
    PartNumber As String
    Description As String
    Material As String
    Document As Inventor.Document
    ComponentOcc As Inventor.ComponentOccurrence
    IBMNum As Long
    AddendumNum As Long
End Type
