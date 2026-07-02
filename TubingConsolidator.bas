Imports Inventor
Imports System.Linq
Imports System.Collections.Generic

Public Class TubingConsolidator

    ' Configuration variables (user-adjustable)
    Private Const TEMPLATE_PATH As String = "C:\Users\hmclinn\OneDrive - winholt.com\Documents\Hunter's Folder\Coding Projects\Autodesk Tools\Tube Template.idw"
    
    ' Template sheet configuration
    Private templateSheets As New Dictionary(Of Integer, String) From {
        {3, "3 Tubes"},
        {4, "4 Tubes"},
        {6, "6 tubes"}
    }

    ' Global objects
    Private invApp As Inventor.Application
    Private invDoc As DrawingDocument
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

            ' Step 3: Create drawing sheets for each IBM's tubes using template
            CreateTubingSheetsByTemplate()

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
    ' STEP 2 & 3: CREATE TUBING SHEETS USING TEMPLATE
    ' ====================================================================
    Private Sub CreateTubingSheetsByTemplate()
        Try
            ' Open template document
            Dim templateDoc As DrawingDocument = invApp.Documents.Open(TEMPLATE_PATH, False)

            If templateDoc Is Nothing Then
                MsgBox("Could not open template file: " & TEMPLATE_PATH, "Error")
                Exit Sub
            End If

            ' Iterate through each IBM
            For Each ibmKey In tubesByIBM.Keys
                Dim ibmNum As Long = ibmKey
                Dim subAsmDict As Dictionary(Of String, List(Of TubeData)) = tubesByIBM(ibmKey)

                ' Iterate through each sub-assembly within this IBM
                For Each subAsmKey In subAsmDict.Keys
                    Dim tubeColl As List(Of TubeData) = subAsmDict(subAsmKey)

                    ' Sort tubes by part number
                    SortTubesByPartNumber(tubeColl)

                    ' Plan the sheet layout based on tube count
                    Dim sheetPlan As List(Of SheetConfig) = PlanSheetLayout(tubeColl.Count)

                    ' Create sheets according to plan
                    Dim tubeIndex As Integer = 0
                    For sheetIdx As Integer = 0 To sheetPlan.Count - 1
                        Dim config As SheetConfig = sheetPlan(sheetIdx)
                        Dim startIdx As Integer = tubeIndex
                        Dim endIdx As Integer = tubeIndex + config.TubeCount - 1

                        ' Create sheet
                        CopyAndPopulateTemplateSheet(templateDoc, tubeColl, startIdx, endIdx, config.TemplateName, sheetIdx + 1, sheetPlan.Count)

                        tubeIndex += config.TubeCount
                    Next
                Next
            Next

            ' Close template without saving
            templateDoc.Close(False)

        Catch ex As Exception
            MsgBox("Error creating sheets from template: " & ex.Message, "Error")
        End Try
    End Sub

    ' ====================================================================
    ' PLAN SHEET LAYOUT BASED ON TUBE COUNT
    ' ====================================================================
    Private Function PlanSheetLayout(tubeCount As Integer) As List(Of SheetConfig)
        Dim plan As New List(Of SheetConfig)

        Select Case tubeCount
            Case 1
                ' Single tube - no template needed, skip entirely
                ' Return empty list to skip
                Return plan

            Case 2, 3
                ' 2-3 tubes - use 3 tube template
                plan.Add(New SheetConfig With {.TubeCount = tubeCount, .TemplateName = "3 Tubes"})

            Case 4
                ' 4 tubes - use 4 tube template
                plan.Add(New SheetConfig With {.TubeCount = 4, .TemplateName = "4 Tubes"})

            Case 5
                ' 5 tubes - split between 3 and 4 tube templates (more even)
                plan.Add(New SheetConfig With {.TubeCount = 3, .TemplateName = "3 Tubes"})
                plan.Add(New SheetConfig With {.TubeCount = 2, .TemplateName = "3 Tubes"})

            Case 6
                ' 6 tubes - use 6 tube template
                plan.Add(New SheetConfig With {.TubeCount = 6, .TemplateName = "6 tubes"})

            Case 7
                ' 7 tubes - split between 4 and 3 tube templates
                plan.Add(New SheetConfig With {.TubeCount = 4, .TemplateName = "4 Tubes"})
                plan.Add(New SheetConfig With {.TubeCount = 3, .TemplateName = "3 Tubes"})

            Case 8
                ' 8 tubes - use two 4 tube templates
                plan.Add(New SheetConfig With {.TubeCount = 4, .TemplateName = "4 Tubes"})
                plan.Add(New SheetConfig With {.TubeCount = 4, .TemplateName = "4 Tubes"})

            Case 9
                ' 9 tubes - use three 3 tube templates
                plan.Add(New SheetConfig With {.TubeCount = 3, .TemplateName = "3 Tubes"})
                plan.Add(New SheetConfig With {.TubeCount = 3, .TemplateName = "3 Tubes"})
                plan.Add(New SheetConfig With {.TubeCount = 3, .TemplateName = "3 Tubes"})

            Case 10
                ' 10 tubes - use 6 tube and 4 tube templates
                plan.Add(New SheetConfig With {.TubeCount = 6, .TemplateName = "6 tubes"})
                plan.Add(New SheetConfig With {.TubeCount = 4, .TemplateName = "4 Tubes"})

            Case 11
                ' 11 tubes - use 6, 4, and split extra (6 + 3 + 2)
                plan.Add(New SheetConfig With {.TubeCount = 6, .TemplateName = "6 tubes"})
                plan.Add(New SheetConfig With {.TubeCount = 4, .TemplateName = "4 Tubes"})
                plan.Add(New SheetConfig With {.TubeCount = 1, .TemplateName = "3 Tubes"})

            Case 12
                ' 12 tubes - use two 6 tube templates
                plan.Add(New SheetConfig With {.TubeCount = 6, .TemplateName = "6 tubes"})
                plan.Add(New SheetConfig With {.TubeCount = 6, .TemplateName = "6 tubes"})

            Case Else
                ' More than 12 tubes - distribute evenly
                Dim remaining As Integer = tubeCount
                While remaining > 0
                    If remaining >= 6 Then
                        plan.Add(New SheetConfig With {.TubeCount = 6, .TemplateName = "6 tubes"})
                        remaining -= 6
                    ElseIf remaining >= 4 Then
                        plan.Add(New SheetConfig With {.TubeCount = 4, .TemplateName = "4 Tubes"})
                        remaining -= 4
                    Else
                        plan.Add(New SheetConfig With {.TubeCount = remaining, .TemplateName = "3 Tubes"})
                        remaining = 0
                    End If
                End While
        End Select

        Return plan
    End Function

    ' ====================================================================
    ' HELPER: SORT TUBES BY PART NUMBER
    ' ====================================================================
    Private Sub SortTubesByPartNumber(tubeColl As List(Of TubeData))
        ' Sort by addendum number
        tubeColl.Sort(Function(a, b) a.AddendumNum.CompareTo(b.AddendumNum))
    End Sub

    ' ====================================================================
    ' COPY AND POPULATE TEMPLATE SHEET
    ' ====================================================================
    Private Sub CopyAndPopulateTemplateSheet(templateDoc As DrawingDocument, tubeColl As List(Of TubeData), startIdx As Integer, endIdx As Integer, templateSheetName As String, sheetNum As Integer, totalSheets As Integer)
        Try
            ' Get template sheet
            Dim templateSheet As Sheet = Nothing
            For i As Integer = 1 To templateDoc.Sheets.Count
                If templateDoc.Sheets(i).Name = templateSheetName Then
                    templateSheet = templateDoc.Sheets(i)
                    Exit For
                End If
            Next

            If templateSheet Is Nothing Then
                MsgBox("Template sheet '" & templateSheetName & "' not found.", "Error")
                Exit Sub
            End If

            ' Copy the template sheet to current document
            Dim newSheet As Sheet = invDoc.Sheets.Add()
            
            ' Create sheet name
            Dim sheetName As String
            If totalSheets > 1 Then
                sheetName = "LEGS & SPREADERS (" & sheetNum & "/" & totalSheets & ")"
            Else
                sheetName = "LEGS & SPREADERS"
            End If
            newSheet.Name = sheetName

            ' Copy all views from template sheet to new sheet
            Dim tubeIndex As Integer = 0
            For i As Integer = 1 To templateSheet.DrawingViews.Count
                Dim templateView As DrawingView = templateSheet.DrawingViews(i)
                Dim viewName As String = templateView.Name

                ' Check if this is a base view (VIEW1, VIEW2, etc.) or projection (VIEW1A, VIEW2A, etc.)
                If Not viewName.EndsWith("A") Then
                    ' This is a base view - replace with tubing model
                    If startIdx + tubeIndex <= endIdx Then
                        Dim tubeData As TubeData = tubeColl(startIdx + tubeIndex)

                        ' Get the position and scale from template view
                        Dim position As Point2d = templateView.Position
                        Dim scale As Double = templateView.Scale

                        ' Add new base view with tubing model
                        Dim newView As DrawingView = newSheet.DrawingViews.AddBaseView(tubeData.Document, position, scale, ViewOrientationTypeEnum.kRightViewOrientation, DrawingViewStyleEnum.kHiddenLineDrawingViewStyle)
                        newView.Name = viewName

                        tubeIndex += 1
                    End If
                Else
                    ' This is a projection view (VIEW1A, VIEW2A, etc.)
                    ' Find the corresponding base view that was just added
                    Dim baseViewName As String = viewName.Substring(0, viewName.Length - 1) ' Remove the 'A'
                    Dim baseView As DrawingView = Nothing

                    ' Find the base view in the new sheet
                    For j As Integer = 1 To newSheet.DrawingViews.Count
                        If newSheet.DrawingViews(j).Name = baseViewName Then
                            baseView = newSheet.DrawingViews(j)
                            Exit For
                        End If
                    Next

                    If baseView IsNot Nothing Then
                        ' Get position from template projection
                        Dim position As Point2d = templateView.Position

                        ' Add projection view based on the base view
                        Dim projView As DrawingView = newSheet.DrawingViews.AddProjectedView(baseView, position, DrawingViewStyleEnum.kHiddenLineDrawingViewStyle, baseView.Scale)
                        projView.Name = viewName
                    End If
                End If
            Next

            ' Copy all sketches and annotations from template sheet (for dimensions, break lines, etc.)
            CopySketchesAndAnnotations(templateSheet, newSheet)

        Catch ex As Exception
            MsgBox("Error populating template sheet: " & ex.Message, "Error")
        End Try
    End Sub

    ' ====================================================================
    ' COPY SKETCHES AND ANNOTATIONS
    ' ====================================================================
    Private Sub CopySketchesAndAnnotations(sourceSheet As Sheet, targetSheet As Sheet)
        Try
            ' Copy all sketches from source to target
            For i As Integer = 1 To sourceSheet.Sketches.Count
                Dim sourceSketch As PlanarSketch = sourceSheet.Sketches(i)
                ' Clone the sketch to the target sheet
                sourceSketch.CopyToClipboard()
                targetSheet.PasteSpecial()
            Next

        Catch ex As Exception
            ' If sketch copy fails, continue without it
            ' The dimensions and break lines from template will be preserved through view properties
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
' SHEET CONFIGURATION CLASS
' ====================================================================
Public Class SheetConfig
    Public TubeCount As Integer
    Public TemplateName As String
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
