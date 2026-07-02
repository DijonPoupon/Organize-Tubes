Imports Inventor
Imports System.Linq
Imports System.Collections.Generic

Public Class TubingConsolidator

    ' Configuration variables (user-adjustable)
    Private VAR_BASE_Y_POS As Double = 9.0
    Private VAR_TUBE_SPACING As Double = 2.2
    Private VAR_BASE_X_POS As Double = 0.75
    Private VAR_CROSS_SECTION_OFFSET_X As Double = 0
    Private VAR_PROJECTION_OFFSET_X As Double = 1.2
    Private VAR_PARTS_LIST_OFFSET_X As Double = 1.8

    ' Global objects
    Private invApp As Inventor.Application
    Private invDoc As DrawingDocument
    Private invSheet As Sheet
    Private tubesByIBM As New Dictionary(Of Long, Dictionary(Of String, List(Of TubeData)))
    Private processedDocs As New HashSet(Of String)

    ' ====================================================================
    ' MAIN ENTRY POINT
    ' ====================================================================
    Public Sub Main()
        Try
            ' Get Inventor application and document
            invApp = ThisApplication
            invDoc = ThisDrawing.Document

            ' Validate document type
            If invDoc Is Nothing Then
                MsgBox("Please open a Drawing document first.", "Error")
                Exit Sub
            End If

            ' Step 1: Get the main assembly from the first sheet
            Dim mainAssemblyDoc As AssemblyDocument = GetAssemblyFromFirstSheet()
            
            If mainAssemblyDoc Is Nothing Then
                MsgBox("Could not find assembly on first sheet.", "Error")
                Exit Sub
            End If

            ' Step 2: Collect tubing parts organized by IBM and sub-assembly
            CollectTubingParts(mainAssemblyDoc)

            ' Validate that tubes were found
            If tubesByIBM.Count = 0 Then
                MsgBox("No tubing parts found (material codes 351-369).", "Error")
                Exit Sub
            End If

            ' Step 3: Create drawing sheets for each IBM's tubes
            CreateTubingSheetsByIBM()

            MsgBox("Tubing consolidation complete! Created " & GetTotalTubingSheets() & " tubing sheet(s).", "Success")

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, "Error")
        End Try
    End Sub

    ' ====================================================================
    ' HELPER: GET ASSEMBLY FROM FIRST SHEET
    ' ====================================================================
    Private Function GetAssemblyFromFirstSheet() As AssemblyDocument
        Try
            Dim sheet As Sheet = invDoc.Sheets(1)
            invSheet = sheet

            ' Loop through views on the first sheet to find an assembly view
            For i As Integer = 1 To sheet.DrawingViews.Count
                Dim view As DrawingView = sheet.DrawingViews(i)
                
                If view.ReferencedDocumentDescriptor.ReferencedDocument.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                    Return view.ReferencedDocumentDescriptor.ReferencedDocument
                End If
            Next

            Return Nothing

        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    ' ====================================================================
    ' STEP 1: COLLECT TUBING PARTS ORGANIZED BY IBM
    ' ====================================================================
    Private Sub CollectTubingParts(asmDoc As AssemblyDocument)
        Try
            Dim asmDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition

            ' Loop through all occurrences in the assembly
            For i As Integer = 1 To asmDef.Occurrences.Count
                Dim occ As ComponentOccurrence = asmDef.Occurrences(i)

                ' Process recursively if this is a sub-assembly
                If occ.Definition.Document.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                    If Not processedDocs.Contains(occ.Definition.Document.FullFileName) Then
                        processedDocs.Add(occ.Definition.Document.FullFileName)
                        CollectTubingParts(occ.Definition.Document)
                    End If
                Else
                    ' This is a part - check if it's a valid tubing part
                    Dim fileName As String = GetFileNameWithoutExtension(occ.Definition.Document.FullFileName)

                    ' Get IBM and Addendum numbers
                    Dim ibmNum As Long = GetIBMNum(fileName)
                    Dim addendumNum As Long = GetAddendumNum(fileName)

                    ' Validate this is a valid part (has both IBM and Addendum)
                    If ibmNum > 0 And addendumNum >= 0 Then
                        ' Check material code
                        Dim materialCode As String = GetIProperty(occ.Definition.Document, "Physical", "Material")

                        ' Check if material code is tubing (351-369)
                        If materialCode.Length >= 3 Then
                            Dim materialNum As Integer = CInt(materialCode.Substring(0, 3))
                            If materialNum >= 351 And materialNum <= 369 Then
                                ' Create tube data structure
                                Dim tubeData As New TubeData With {
                                    .PartName = occ.Name,
                                    .PartNumber = fileName,
                                    .Description = GetIProperty(occ.Definition.Document, "Summary", "Title"),
                                    .Material = materialCode,
                                    .Document = occ.Definition.Document,
                                    .ComponentOcc = occ,
                                    .IBMNum = ibmNum,
                                    .AddendumNum = addendumNum
                                }

                                ' Create IBM dictionary if it doesn't exist
                                If Not tubesByIBM.ContainsKey(ibmNum) Then
                                    tubesByIBM.Add(ibmNum, New Dictionary(Of String, List(Of TubeData)))
                                End If

                                Dim subAsmDict As Dictionary(Of String, List(Of TubeData)) = tubesByIBM(ibmNum)
                                Dim subAsmKey As String = "Main"

                                ' Create sub-assembly collection if it doesn't exist
                                If Not subAsmDict.ContainsKey(subAsmKey) Then
                                    subAsmDict.Add(subAsmKey, New List(Of TubeData))
                                End If

                                ' Add to appropriate collection
                                subAsmDict(subAsmKey).Add(tubeData)
                            End If
                        End If
                    End If
                End If
            Next

        Catch ex As Exception
            ' Silently continue on error
        End Try
    End Sub

    ' ====================================================================
    ' HELPER: GET FILE NAME WITHOUT EXTENSION
    ' ====================================================================
    Private Function GetFileNameWithoutExtension(fullPath As String) As String
        Dim fileName As String = System.IO.Path.GetFileNameWithoutExtension(fullPath)
        Return fileName
    End Function

    ' ====================================================================
    ' HELPER: GET IBM NUMBER
    ' ====================================================================
    Private Function GetIBMNum(fileName As String) As Long
        Try
            Dim parts As String() = fileName.Split(" "c)
            If parts.Length < 1 Then
                Return -1
            End If

            Dim splitDash As String() = parts(0).Split("-"c)
            If splitDash.Length <> 2 Then
                Return -1
            End If

            ' Validate IBM number
            If splitDash(0).Length < 6 Then
                Return -1
            End If

            Dim ibmNum As Long
            If Long.TryParse(splitDash(0), ibmNum) Then
                Return ibmNum
            End If

            Return -1

        Catch ex As Exception
            Return -1
        End Try
    End Function

    ' ====================================================================
    ' HELPER: GET ADDENDUM NUMBER
    ' ====================================================================
    Private Function GetAddendumNum(fileName As String) As Long
        Try
            Dim parts As String() = fileName.Split(" "c)
            If parts.Length < 1 Then
                Return -1
            End If

            Dim splitDash As String() = parts(0).Split("-"c)
            If splitDash.Length <> 2 Then
                Return -1
            End If

            ' Validate IBM part
            If splitDash(0).Length < 6 Then
                Return -1
            End If

            Dim ibmNum As Long
            If Not Long.TryParse(splitDash(0), ibmNum) Then
                Return -1
            End If

            ' Try to parse addendum number
            Dim addendumNum As Long
            If Long.TryParse(splitDash(1), addendumNum) Then
                Return addendumNum
            End If

            Return -1

        Catch ex As Exception
            Return -1
        End Try
    End Function

    ' ====================================================================
    ' HELPER: GET IPROPERTY VALUE
    ' ====================================================================
    Private Function GetIProperty(doc As Document, section As String, propertyName As String) As String
        Try
            Dim propSet = doc.PropertySets.Item(section)
            If propSet Is Nothing Then
                Return ""
            End If

            Dim prop = propSet.Item(propertyName)
            If prop Is Nothing Then
                Return ""
            End If

            Return prop.Value.ToString()

        Catch ex As Exception
            Return ""
        End Try
    End Function

    ' ====================================================================
    ' STEP 2 & 3: CREATE TUBING SHEETS FOR EACH IBM
    ' ====================================================================
    Private Sub CreateTubingSheetsByIBM()
        Dim tubesPerSheet As Integer = 6

        ' Iterate through each IBM
        For Each ibmKey In tubesByIBM.Keys
            Dim ibmNum As Long = ibmKey
            Dim subAsmDict As Dictionary(Of String, List(Of TubeData)) = tubesByIBM(ibmKey)

            ' Iterate through each sub-assembly within this IBM
            For Each subAsmKey In subAsmDict.Keys
                Dim tubeColl As List(Of TubeData) = subAsmDict(subAsmKey)

                ' Sort tubes by part number
                SortTubesByPartNumber(tubeColl)

                ' Calculate number of sheets needed for this sub-assembly
                Dim sheetCount As Integer = CInt(Math.Ceiling(tubeColl.Count / CDbl(tubesPerSheet)))

                ' Create sheets for this sub-assembly
                For currentSheet As Integer = 1 To sheetCount
                    Dim startIdx As Integer = (currentSheet - 1) * tubesPerSheet
                    Dim endIdx As Integer = Math.Min(startIdx + tubesPerSheet - 1, tubeColl.Count - 1)

                    ' Create new sheet
                    Dim newSheet As Sheet = CreateNewTubingSheet(ibmNum, subAsmKey, currentSheet, sheetCount)

                    ' Determine scale based on tube count
                    Dim tubeCountForScale As Integer = endIdx - startIdx + 1
                    Dim scale As Double = DetermineScale(tubeCountForScale)

                    ' Place tubes on sheet
                    PlaceTubesOnSheet(newSheet, tubeColl, startIdx, endIdx, scale)
                Next
            Next
        Next
    End Sub

    ' ====================================================================
    ' HELPER: SORT TUBES BY PART NUMBER
    ' ====================================================================
    Private Sub SortTubesByPartNumber(tubeColl As List(Of TubeData))
        ' Sort by addendum number
        tubeColl.Sort(Function(a, b) a.AddendumNum.CompareTo(b.AddendumNum))
    End Sub

    ' ====================================================================
    ' HELPER: CREATE NEW TUBING SHEET
    ' ====================================================================
    Private Function CreateNewTubingSheet(ibmNum As Long, subAsmName As String, sheetNum As Integer, totalSheets As Integer) As Sheet
        Try
            ' Create sheet name
            Dim sheetName As String
            If totalSheets > 1 Then
                sheetName = "LEGS & SPREADERS - " & ibmNum & " (" & sheetNum & "/" & totalSheets & ")"
            Else
                sheetName = "LEGS & SPREADERS - " & ibmNum
            End If

            ' Create new sheet
            Dim newSheet As Sheet = invDoc.Sheets.Add(DrawingSheetSizeEnum.kADrawingSheetSize)
            newSheet.Name = sheetName

            ' Try to apply WINHOLT title block if available
            Try
                Dim titleBlockDef As TitleBlockDefinition = invDoc.TitleBlockDefinitions.Item("WINHOLT")

                If newSheet.TitleBlock IsNot Nothing Then
                    newSheet.TitleBlock.Delete()
                End If

                Dim sPromptStrings(2) As String
                sPromptStrings(0) = ""
                sPromptStrings(1) = "LEGS & SPREADERS"
                sPromptStrings(2) = ""

                newSheet.AddTitleBlock(titleBlockDef, , sPromptStrings)
            Catch ex As Exception
                ' Title block not available, continue without it
            End Try

            Return newSheet

        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    ' ====================================================================
    ' HELPER: DETERMINE SCALE BASED ON TUBE COUNT
    ' ====================================================================
    Private Function DetermineScale(tubeCount As Integer) As Double
        Select Case tubeCount
            Case 2, 3
                Return 0.25
            Case 4, 5, 6
                Return 0.1875
            Case Else
                Return 0.1875
        End Select
    End Function

    ' ====================================================================
    ' STEP 4: PLACE TUBES ON SHEET
    ' ====================================================================
    Private Sub PlaceTubesOnSheet(sheet As Sheet, tubeColl As List(Of TubeData), startIdx As Integer, endIdx As Integer, scale As Double)
        Try
            Dim tg As TransientGeometry = invApp.TransientGeometry
            Dim crossSectionX As Double = VAR_BASE_X_POS + VAR_CROSS_SECTION_OFFSET_X

            ' Place each tube
            For i As Integer = startIdx To endIdx
                Dim tubeData As TubeData = tubeColl(i)
                Dim yPos As Double = VAR_BASE_Y_POS - ((i - startIdx) * VAR_TUBE_SPACING)

                ' Create cross-section view
                Try
                    Dim view As DrawingView = PlaceCrossSectionView(sheet, tubeData, crossSectionX, yPos, scale, tg)

                    If view IsNot Nothing Then
                        ' Create projection view with length
                        Try
                            Dim projView As DrawingView = PlaceLengthProjectionView(sheet, view, tubeData, scale, tg)

                            If projView IsNot Nothing Then
                                ' Add length dimension to projection view
                                AddLengthDimension(sheet, projView, tubeData, tg)

                                ' Add parts list box
                                AddPartsList(sheet, tubeData, projView, tg)

                                ' Check and add break line if necessary
                                If NeedsBreakLine(tubeData, scale, sheet) Then
                                    AddBreakLine(sheet, projView)
                                End If
                            End If
                        Catch ex As Exception
                        End Try
                    End If
                Catch ex As Exception
                End Try
            Next

        Catch ex As Exception
        End Try
    End Sub

    ' ====================================================================
    ' HELPER: PLACE CROSS-SECTION VIEW
    ' ====================================================================
    Private Function PlaceCrossSectionView(sheet As Sheet, tubeData As TubeData, xPos As Double, yPos As Double, scale As Double, tg As TransientGeometry) As DrawingView
        Try
            Dim placement As Point2d = tg.CreatePoint2d(xPos, yPos)
            Dim view As DrawingView = sheet.DrawingViews.AddBaseView(tubeData.Document, placement, scale, ViewOrientationTypeEnum.kRightViewOrientation, DrawingViewStyleEnum.kHiddenLineDrawingViewStyle)
            Return view

        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    ' ====================================================================
    ' HELPER: PLACE LENGTH PROJECTION VIEW
    ' ====================================================================
    Private Function PlaceLengthProjectionView(sheet As Sheet, baseView As DrawingView, tubeData As TubeData, scale As Double, tg As TransientGeometry) As DrawingView
        Try
            Dim projXPos As Double = baseView.Position.X + VAR_PROJECTION_OFFSET_X
            Dim projPlacement As Point2d = tg.CreatePoint2d(projXPos, baseView.Position.Y)

            ' Create projection view (top view to show length)
            Dim projView As DrawingView = sheet.DrawingViews.AddProjectedView(baseView, projPlacement, DrawingViewStyleEnum.kHiddenLineDrawingViewStyle, scale)
            Return projView

        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    ' ====================================================================
    ' HELPER: ADD LENGTH DIMENSION
    ' ====================================================================
    Private Sub AddLengthDimension(sheet As Sheet, projView As DrawingView, tubeData As TubeData, tg As TransientGeometry)
        Try
            ' Placeholder for dimension logic
            ' This would need to extract actual tube length and add dimension

        Catch ex As Exception
        End Try
    End Sub

    ' ====================================================================
    ' HELPER: ADD PARTS LIST
    ' ====================================================================
    Private Sub AddPartsList(sheet As Sheet, tubeData As TubeData, refView As DrawingView, tg As TransientGeometry)
        Try
            Dim xPos As Double = refView.Position.X + VAR_PARTS_LIST_OFFSET_X
            Dim yPos As Double = refView.Position.Y
            Dim placement As Point2d = tg.CreatePoint2d(xPos, yPos)

            ' Create text box with parts list information
            Dim partsListText As String = "PART NUMBER" & vbCrLf & tubeData.PartNumber & vbCrLf & vbCrLf & _
                                         "DESCRIPTION" & vbCrLf & tubeData.Description & vbCrLf & vbCrLf & _
                                         "MATERIAL" & vbCrLf & tubeData.Material

            ' Note: Text box creation would need proper annotation implementation
            ' This is a placeholder for actual parts list creation

        Catch ex As Exception
        End Try
    End Sub

    ' ====================================================================
    ' HELPER: CHECK IF BREAK LINE NEEDED
    ' ====================================================================
    Private Function NeedsBreakLine(tubeData As TubeData, scale As Double, sheet As Sheet) As Boolean
        Try
            ' Standard sheet height is 11" (A size)
            Dim sheetHeight As Double = 11.0
            ' 1/4 of sheet length at current scale
            Dim maxAllowedLength As Double = (sheetHeight / 4.0) / scale

            ' Placeholder for actual break line logic
            Return False

        Catch ex As Exception
            Return False
        End Try
    End Function

    ' ====================================================================
    ' HELPER: ADD BREAK LINE
    ' ====================================================================
    Private Sub AddBreakLine(sheet As Sheet, projView As DrawingView)
        Try
            ' Placeholder for break line implementation

        Catch ex As Exception
        End Try
    End Sub

    ' ====================================================================
    ' HELPER: COUNT TOTAL TUBING SHEETS
    ' ====================================================================
    Private Function GetTotalTubingSheets() As Integer
        Dim count As Integer = 0
        For i As Integer = 1 To invDoc.Sheets.Count
            If invDoc.Sheets(i).Name.Contains("LEGS & SPREADERS") Then
                count += 1
            End If
        Next
        Return count
    End Function

End Class

' ====================================================================
' DATA STRUCTURE: TUBE DATA
' ====================================================================
Public Class TubeData
    Public PartName As String
    Public PartNumber As String
    Public Description As String
    Public Material As String
    Public Document As Document
    Public ComponentOcc As ComponentOccurrence
    Public IBMNum As Long
    Public AddendumNum As Long
End Class
