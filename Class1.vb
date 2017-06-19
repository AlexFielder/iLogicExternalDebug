Option Explicit On
Imports System.IO
Imports System.Linq
Imports System.Windows.Forms
Imports Inventor
Imports Autodesk.iLogic.Interfaces
Imports iLogicExternalDebug
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices
Imports log4net
Imports System.Text.RegularExpressions

Public Class ExtClass
    '#Region "Properties"
    ''' <summary>
    ''' useful for passing the inventor application to this .dll
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared m_inventorApplication As Inventor.Application
    Private excelapp As Excel.Application
    Private workBook As Workbook = Nothing
    Private usedRange As Excel.Range
    Private m_windowHandle As IntPtr
    Private m_processID As IntPtr
    Private thisAssembly As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
    Private thisAssemblyPath As String = String.Empty
    Private logHelper As Log4NetFileHelper = New Log4NetFileHelper()
    Private Shared ReadOnly log As ILog = LogManager.GetLogger(GetType(ExtClass))
    Private Const ExcelWindowCaption As String = "Running From CAFE Tool O_o"

    Public viewStateList As List(Of viewstate) = New List(Of viewstate)
    Public Shared dwgDoc As DrawingDocument = Nothing
    Public ActiveSht As Sheet = Nothing
    Public oDrawingViews As DrawingViews = Nothing
    'footwalks
    Public AssyDoc As AssemblyDocument
    Public AssyDef As AssemblyComponentDefinition


    ''' <summary>
    ''' the Inventor object populated by m_inventorApplication
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Property ThisApplication() As Inventor.Application
        Get
            ThisApplication = m_inventorApplication
        End Get

        Set(ByVal Value As Inventor.Application)
            m_inventorApplication = Value
        End Set
    End Property

    Private m_DocToUpdate As ICadDoc
    Public Property DocToUpdate() As ICadDoc
        Get
            Return m_DocToUpdate
        End Get
        Set(ByVal value As ICadDoc)
            m_DocToUpdate = value
        End Set
    End Property

#Region "Patterning Footwalks"
    '''Needs to use Module_Pattern_QTY for driving the number, Module_Spacing for the distance between
    ''' and the Z axis for direction.
    '''
    Public Sub PatternFootWalk(OccName As String, NumOccs As Integer, OffsetDistance As Double, PatternName As String, FolderName As String)
        Dim CompOccs As ComponentOccurrences = AssyDef.Occurrences
        Dim newPatternOcc As RectangularOccurrencePattern
        Dim compOcc As ComponentOccurrence = CompOccs.ItemByName(OccName)
        Dim objCol As ObjectCollection = ThisApplication.TransientObjects.CreateObjectCollection

        'base work axes - don't need all three but useful to demonstrate
        Dim XAxis As WorkAxis
        Dim YAxis As WorkAxis
        Dim Zaxis As WorkAxis
        With AssyDef
            XAxis = .WorkAxes(1)
            YAxis = .WorkAxes(2)
            Zaxis = .WorkAxes(3)
        End With

        objCol.Add(compOcc)
        newPatternOcc = AssyDef.OccurrencePatterns.AddRectangularPattern(objCol,
                                                                        Zaxis,
                                                                        False,
                                                                        OffsetDistance / 10,
                                                                        NumOccs)
        newPatternOcc.Name = PatternName

        'AddNewPatternToFolder(FolderName, newPatternOcc)
        '	Component.IsActive("Footwalk Inter RH Pattern") = False
        '	Component.IsActive("Footwalk Inter LH Pattern") = False

    End Sub

    Sub AddNewPatternToFolder(Foldername As String, occ As RectangularOccurrencePattern)
        ' get the model browser pane
        Dim oPane As BrowserPane
        oPane = assyDoc.BrowserPanes.Item("Model")

        ' Create a Browser node object from an existing object
        Dim oNode As BrowserNode
        oNode = oPane.GetBrowserNodeFromObject(occ)



        ' Add the node to the extra folder
        Dim browserFolder As BrowserFolder = (From a As BrowserFolder In oPane.TopNode.BrowserFolders
                                              Where a.Name = Foldername
                                              Select a).FirstOrDefault()
        Dim maleNode As BrowserNode = (From node As BrowserNode In browserFolder.BrowserNode.BrowserNodes
                                       Where node.BrowserNodeDefinition.Label.StartsWith("Footwalk Male")
                                       Select node).FirstOrDefault()
        'browserFolder.Add(oNode, maleNode, False)
        'doesn't f*$£ing work!
        'oPane.TopNode.BrowserFolders.Item(Foldername).Add(oNode)
    End Sub

    Public Sub ActuallyDeletePattern(PatternName As String)
        '	Try
        Dim CompOccs As ComponentOccurrences = AssyDef.Occurrences
        'Dim PatternOccToDelete As RectangularOccurrencePattern = CompOccs.ItemByName(PatternName)
        ' Dim PatternOccToDelete As ComponentOccurrence = CompOccs.ItemByName(PatternName)

        ' get the model browser pane
        Dim oPane As BrowserPane
        oPane = AssyDoc.BrowserPanes.Item("Model")
        Dim nodeTodelete As BrowserNode = (From node As BrowserNode In oPane.TopNode.BrowserNodes
                                           Where node.BrowserNodeDefinition.Label = PatternName
                                           Select node).FirstOrDefault()
        If Not nodeTodelete Is Nothing Then
            Dim PatternOccToDelete As RectangularOccurrencePattern = nodeTodelete.NativeObject
            PatternOccToDelete.Delete()
        End If
    End Sub

#End Region

#Region "Align drawing views"

    Public Sub BeginAlignDrawingviews()
        thisAssemblyPath = System.IO.Path.GetDirectoryName(thisAssembly.Location)
        logHelper.Init()
        logHelper.AddConsoleLogging()
        'the next line works but we want a rolling log.
        logHelper.AddFileLogging(System.IO.Path.Combine(thisAssemblyPath, "GraitecExtensionsServer.log"))
        logHelper.AddFileLogging("C:\Logs\MyLogFile.txt", log4net.Core.Level.All, True)
        logHelper.AddRollingFileLogging("C:\Logs\RollingFileLog.txt", log4net.Core.Level.All, True)
        log.Debug("Loading iLogic External Debug for align drawing views")
        aligndrawingviews()

    End Sub
    Sub aligndrawingviews()
        'set up logging
        logHelper.Init()
        logHelper.AddRollingFileLogging("C:\Logs\RollingFileLog.txt", log4net.Core.Level.All, True)
        log.Debug("logging from iLogic, who'd have thought it!?")
        If Not TypeOf (ThisApplication.ActiveDocument) Is DrawingDocument Then
            log.Error("Drawing not active!")
            Exit Sub
        End If
        dwgDoc = ThisApplication.ActiveDocument
        Dim tr As Transaction = ThisApplication.TransactionManager.StartTransaction(dwgDoc, "Align Drawing Views")
        Try
            Dim viewState As viewstate = Nothing
            ActiveSht = dwgDoc.ActiveSheet
            oDrawingViews = ActiveSht.DrawingViews

            'capture initialview states to re-enable them later.
            For Each dwgView As DrawingView In odrawingviews
                viewState = New viewstate
                viewState.View = dwgView
                'viewstate.curves = dwgview.drawingcurves
                viewState.viewenabled = dwgView.Suppressed
                viewstatelist.add(viewState)
            Next
            log.Debug("captured viewstates")
            Dim selectedDrawingView1 As DrawingView = ThisApplication.CommandManager.Pick(SelectionFilterEnum.kDrawingViewFilter, "Select a drawing view.")
            'turn off other competing views
            SetViewVisibility(selectedDrawingView1, False)
            Dim selectedCurve1 As DrawingCurveSegment = ThisApplication.CommandManager.Pick(SelectionFilterEnum.kDrawingCurveSegmentFilter, "Select a point within the chosen view.")
            'turn all drawing views on
            SetViewVisibility(Nothing, True)
            Dim selectedDrawingView2 As DrawingView = ThisApplication.CommandManager.Pick(SelectionFilterEnum.kDrawingViewFilter, "Select a drawing view.")
            'turn off other competing views
            SetViewVisibility(selectedDrawingView2, False)
            Dim selectedCurve2 As DrawingCurveSegment = ThisApplication.CommandManager.Pick(SelectionFilterEnum.kDrawingCurveSegmentFilter, "Select a point within the chosen view.")

            SetViewVisibility(Nothing, True)

            'do some more fancy stuff here such as: 
            'display start /end points
            'snap selected points together

            Dim pointToSnap1 As Point2d = DisplayEndPoints(selectedDrawingView1, selectedCurve1)

            Dim pointToSnap2 As Point2d = DisplayEndPoints(selectedDrawingView2, selectedCurve2)



            'not required:
            'Dim highlight1 As HighlightSet = ThisApplication.ActiveDocument.CreateHighlightSet
            'highlight1.AddItem(selectedCurve1.StartPoint)
            'highlight1.AddItem(selectedCurve1.EndPoint)

            Dim tmpPoint As Point2d = selectedDrawingView2.Position 'ThisApplication.TransientGeometry.CreatePoint2d(pointToSnap1.X, pointToSnap1.Y)
            'Dim dwgViewMatrix As Matrix = selectedDrawingView2.DrawingViewToSheetTransform
            'assume that view 2 is the one we are moving
            Dim VectorX As Double = pointToSnap1.X - pointToSnap2.X
            Dim VectorY As Double = pointToSnap1.Y - pointToSnap2.Y
            'dwgViewMatrix.SetTranslation(ThisApplication.TransientGeometry.CreateVector2d(VectorX, VectorY))
            Dim newVector As Vector2d = ThisApplication.TransientGeometry.CreateVector2d(VectorX, VectorY)
            tmpPoint.TranslateBy(newVector)
            selectedDrawingView2.Position = tmpPoint


        Catch ex As Exception
            tr.Abort()
            log.Error(ex.Message, ex)
            'reset views
            SetViewVisibility(Nothing, True)
        Finally
            tr.End()
        End Try
    End Sub

    Private Function DisplayEndPoints(selectedDrawingView As DrawingView, selectedCurve As DrawingCurveSegment) As Point2d
        Dim curve As DrawingCurve = selectedCurve.Parent
        Dim transgeom As TransientGeometry = ThisApplication.TransientGeometry
        Dim clientgraphics1 As ClientGraphics = Nothing
        Try
            clientgraphics1 = selectedDrawingView.ClientGraphicsCollection("selectedlineView1")
        Catch ex As Exception
            clientgraphics1 = selectedDrawingView.ClientGraphicsCollection.Add("selectedlineView1")
        End Try

        Dim gfxnode1 As GraphicsNode = clientgraphics1.AddNode(1)
        Dim endtxtgfx As TextGraphics = gfxnode1.AddTextGraphics()
        endtxtgfx.Text = "End point"
        endtxtgfx.Anchor = transgeom.CreatePoint(selectedCurve.EndPoint.X, selectedCurve.EndPoint.Y, 0)

        Dim starttxtgfx As TextGraphics = gfxnode1.AddTextGraphics()
        starttxtgfx.Text = "Start Point"
        starttxtgfx.Anchor = transgeom.CreatePoint(selectedCurve.StartPoint.X, selectedCurve.StartPoint.Y, 0)

        ThisApplication.ActiveView.Update()

        Dim result As String = InputBox("Which point?", "1 is start; 2 is end", "1")
        If result = "1" Then
            Return selectedCurve.StartPoint
        ElseIf result = "2" Then
            Return selectedCurve.EndPoint
        End If
    End Function

    Public Sub SetViewVisibility(viewToKeepVisible As DrawingView, Optional viewVis As Boolean = False)
        'log.Debug("made it to SetViewVisibility")
        Try
            If Not viewToKeepVisible Is Nothing Then
                'MessageBox.Show("view: " & viewToKeepVisible.Name)
                For Each view As DrawingView In oDrawingViews
                    'MessageBox.Show("view: " & view.Name)
                    If Not view Is viewToKeepVisible Then
                        view.Suppressed = True
                        For Each curve As DrawingCurve In view.DrawingCurves
                            view.SetVisibility(curve, viewVis)
                        Next
                    End If
                Next
            End If
            'restore original settings:
            If viewVis = True And viewToKeepVisible Is Nothing Then
                For Each view As DrawingView In oDrawingViews
                    Dim viewtorestore As viewstate = (From a As viewstate In viewStateList
                                                      Where a.view Is view
                                                      Select a).FirstOrDefault()
                    If Not viewtorestore Is Nothing Then
                        view.Suppressed = viewtorestore.viewEnabled
                    End If
                Next
            End If
        Catch ex As Exception
            log.Error(ex.Message, ex)
        End Try
    End Sub

#End Region
#Region "Check iFactory for errors"
    Public Sub CheckiPartTableForErrors()
        Dim oErrorManager As ErrorManager = ThisApplication.ErrorManager

        If TypeOf (ThisApplication.ActiveDocument) Is PartDocument Then
            Dim oDoc As PartDocument = ThisApplication.ActiveDocument
            Dim oiPart As iPartFactory = oDoc.ComponentDefinition.iPartFactory
            Dim oTop As BrowserNode = oDoc.BrowserPanes("Model").TopNode
            Dim bHasErrorOrWarning As Boolean
            Dim i As Integer
            'InventorVb.DocumentUpdate()
            ThisApplication.SilentOperation = True
            For i = 1 To oiPart.TableRows.Count 'use first 10 rows only for debugging purposes!
                ' Highlight the 3rd iPart table row which has invalid data
                oTop.BrowserNodes("Table").BrowserNodes.Item(i).DoSelect()

                ' Activate the iPart table row
                Dim oCommand As ControlDefinition = ThisApplication.CommandManager.ControlDefinitions("PartComputeiPartRowCtxCmd")
                oCommand.Execute()

                ThisApplication.SilentOperation = False
                ThisApplication.CommandManager.ControlDefinitions.Item("AppZoomallCmd").Execute()

                If oErrorManager.HasErrors Or oErrorManager.HasWarnings Then
                    MessageBox.Show(oErrorManager.LastMessage, "Title")
                End If
            Next i
            MessageBox.Show("No errors shown = None found!", "Title")
        ElseIf TypeOf (ThisApplication.ActiveDocument) Is AssemblyDocument Then
            Dim odoc As AssemblyDocument = ThisApplication.ActiveDocument
            Dim iAssy As iAssemblyFactory = odoc.ComponentDefinition.iAssemblyFactory
            Dim oTop As BrowserNode = odoc.BrowserPanes("Model").TopNode
            Dim bHasErrorOrWarning As Boolean
            Dim i As Integer
            'InventorVb.DocumentUpdate()
            ThisApplication.SilentOperation = True
            For rowIndex = 1 To iAssy.TableRows.Count
                oTop.BrowserNodes("Table").BrowserNodes.Item(rowIndex).DoSelect()
                Dim oCommand As ControlDefinition = ThisApplication.CommandManager.ControlDefinitions("PartComputeiPartRowCtxCmd")
                oCommand.Execute()
                ThisApplication.SilentOperation = False
                ThisApplication.CommandManager.ControlDefinitions.Item("AppZoomallCmd").Execute()

                If oErrorManager.HasErrors Or oErrorManager.HasWarnings Then
                    MessageBox.Show(oErrorManager.LastMessage, "Title")
                End If

            Next
        End If
    End Sub
#End Region


#Region "SaveCopyAsFromBrowserNodeNames"

    Public Sub SaveCopyAsFromBrowserNodeNames()
        If TypeOf (ThisApplication.ActiveDocument) Is AssemblyDocument Then
            Dim searchstring As String = InputBox("what are we searching for?", "Search string", "Default Entry")

            If Not searchstring = String.Empty Then
                Dim oDoc As AssemblyDocument = ThisApplication.ActiveDocument
                Dim oPane As BrowserPane = oDoc.BrowserPanes("Model")
                Dim oTopNode As BrowserNode = oPane.TopNode

                Dim nodelist As List(Of String) = New List(Of String)

                nodelist = (From a As BrowserNode In oTopNode.BrowserNodes
                            Let nodedef As BrowserNodeDefinition = a.BrowserNodeDefinition
                            Where nodedef.Label.Contains(searchstring)
                            Select nodedef.Label).ToList()
                ThisApplication.StatusBarText = searchstring

                '	For Each node As browsernode In oTopnode.Browsernodes
                '		Dim nodeDef as browsernodedefinition = node.browsernodedefinition
                '		If nodedef.label.startswith(searchstring) Then
                '			nodelist.add(nodedef.label)
                '			
                '		End If
                '		MessageBox.Show("browser node: " & nodedef.label)
                '	Next
                If nodelist.Count > 0 Then
                    Dim FolderName As String = System.IO.Path.GetDirectoryName(oDoc.FullFileName)
                    Dim filename As String = System.IO.Path.GetFileNameWithoutExtension(oDoc.FullFileName)
                    For Each nodename As String In nodelist
                        Dim newfilename As String = FolderName & "\" & filename & "-" & nodename.Replace("°:1", "") & ".iam"
                        'newfilename.Replace("°:","°-")
                        MessageBox.Show(newfilename)
                        If Not System.IO.File.Exists(newfilename) Then
                            ThisApplication.ActiveDocument.SaveAs(newfilename, False)
                        End If
                    Next
                End If
            End If
        End If
    End Sub

#End Region

#Region "Renumber Item lists"
    Public Sub BeginRenumberItems()
        Dim sw As New Stopwatch()
        sw.Start()
        thisAssemblyPath = System.IO.Path.GetDirectoryName(thisAssembly.Location)
        logHelper.Init()
        logHelper.AddConsoleLogging()
        'the next line works but we want a rolling log.
        logHelper.AddFileLogging(System.IO.Path.Combine(thisAssemblyPath, "GraitecExtensionsServer.log"))
        logHelper.AddFileLogging("C:\Logs\MyLogFile.txt", log4net.Core.Level.All, True)
        logHelper.AddRollingFileLogging("C:\Logs\RollingFileLog.txt", log4net.Core.Level.All, True)
        log.Debug("Loading iLogic External Debug for renumber Items")
        CheckForReuseAndClearExistingAttributes()
        RunRenumberItems()
        sw.Stop()
        Dim timeElapsed As TimeSpan = sw.Elapsed
        MessageBox.Show("Processing took: " & String.Format("{0:00}:{1:00}:{2:00}.{3:00}",
                                                            timeElapsed.Hours,
                                                            timeElapsed.Minutes,
                                                            timeElapsed.Seconds,
                                                            timeElapsed.Milliseconds / 10))
    End Sub
    Private Sub CheckForReuseAndClearExistingAttributes()
        Dim attribSetEnum As AttributeSetsEnumerator = ThisApplication.ActiveDocument.AttributeManager.FindAttributeSets("CCPartNumberSet*")
        'Dim existingAttributes As ObjectCollection =
        '    ThisApplication.ActiveDocument.AttributeManager.FindObjects("CCPartNumberSet*")
        If Not attribSetEnum.Count = 0 Then
            ccBomRowItems = New List(Of BomRowItem)
            For Each attSet As AttributeSet In attribSetEnum
                Dim filename As Attribute = attSet("FileName")
                Dim partno As Attribute = attSet("StandardPartNum")
                Dim existingItem As New BomRowItem() With {
                    .Document = filename.Value,
                    .ItemNo = partno.Value}
                ccBomRowItems.Add(existingItem)
            Next
        End If


        'If Not existingAttributes.Count = 0 Then
        '    ccBomRowItems = New List(Of BomRowItem)
        '    'For i = 1 To existingAttributes.Count
        '    '    If TypeOf (existingAttributes(i)) Is Document Then
        '    Dim thisDoc As Document = existingAttributes(1)

        '    Dim attSet As AttributeSet = thisDoc.AttributeSets.Item(1)
        '    Dim filename As Attribute = attSet("FileName")
        '    Dim partno As Attribute = attSet("StandardPartNum")
        '    Dim existingItem As New BomRowItem() With {
        '        .Document = filename.Value,
        '        .ItemNo = partno.Value}
        '    ccBomRowItems.Add(existingItem)
        '    '    End If
        '    'Next
        'End If
    End Sub
    Private AssySubAssemblies As List(Of AssemblyDocument) = Nothing
    Private itemisedPartsList As List(Of Document) = Nothing
    Private ccPartsList As List(Of Document) = Nothing
    Private ccQuantityList = Nothing
    Private ccBomRowItems As List(Of BomRowItem) = Nothing
    Private ItemNo As Integer = 500
    Private Sub RunRenumberItems()
        Try


            If TypeOf (ThisApplication.ActiveDocument) Is AssemblyDocument Then
                Dim AssyDoc As AssemblyDocument = ThisApplication.ActiveDocument
                itemisedPartsList = (From item As Document In AssyDoc.AllReferencedDocuments
                                     Order By item.FullFileName
                                     Select item).ToList()

                AssySubAssemblies = (From subassyDoc As Document In AssyDoc.AllReferencedDocuments
                                     Where TypeOf (AssyDoc) Is AssemblyDocument
                                     Let selectedDoc As AssemblyDocument = AssyDoc
                                     Where selectedDoc.ComponentDefinition.BOM.StructuredViewEnabled = True
                                     Select selectedDoc).ToList()

                ccPartsList = (From ccDoc As Document In AssyDoc.AllReferencedDocuments
                               Where ccDoc.FullFileName.Contains("Content Center")
                               Let foldername As String = IO.Path.GetDirectoryName(ccDoc.FullFileName)
                               Order By foldername Ascending
                               Select ccDoc).Distinct().ToList()
                ccPartsList.RemoveAll(Function(x As PartDocument) x.ComponentDefinition.BOMStructure = BOMStructureEnum.kReferenceBOMStructure)

                Dim tmpBomRowItems As List(Of BomRowItem) = New List(Of BomRowItem)


                'this looks correct but will only give us one occurrence of each component.
                For Each ccPart As Document In ccPartsList
                    For Each subAssy As AssemblyDocument In AssySubAssemblies
                        Dim assyCompDef As AssemblyComponentDefinition = subAssy.ComponentDefinition
                        Dim structuredBomView As BOMView = assyCompDef.BOM.BOMViews.Item("Structured")
                        Dim ccPartBomRow As BOMRow = (From row As BOMRow In structuredBomView.BOMRows
                                                      Let RowCompDef As ComponentDefinition = row.ComponentDefinitions(1)
                                                      Let thisDoc As Document = RowCompDef.Document
                                                      Where thisDoc.FullFileName = ccPart.FullFileName
                                                      Select row).FirstOrDefault()
                        If Not ccPartBomRow Is Nothing Then
                            Dim tmpitem As New BomRowItem() With {
                            .Document = ccPartBomRow.ComponentDefinitions(1).Document.FullFileName,
                            .ItemNo = ccPartBomRow.ItemNumber,
                            .Material = iProperties.GetorSetStandardiProperty(
                            ccPartBomRow.ComponentDefinitions(1).Document,
                            PropertiesForDesignTrackingPropertiesEnum.kMaterialDesignTrackingProperties, "", ""),
                            .Quantity = ccPartBomRow.TotalQuantity}
                            tmpBomRowItems.Add(tmpitem)
                        End If
                    Next
                Next

                Dim tmpoccurrenceslist As List(Of BomRowItem) = processSubAssys(AssyDef)

                'Dim tmpoccurrenceslist As List(Of BomRowItem) = New List(Of BomRowItem)

                'For Each ccpart As Document In ccPartsList
                '    For Each occ As ComponentOccurrence In AssyDef.Occurrences
                '        Dim occDoc As Document = occ.Definition.Document
                '        If occDoc Is ccpart Then
                '            Dim tmpOccItem As New BomRowItem() With {
                '                .Document = occDoc.FullFileName,
                '                .ItemNo = 0,
                '                .Material = iProperties.GetorSetStandardiProperty(occDoc, PropertiesForDesignTrackingPropertiesEnum.kMaterialDesignTrackingProperties, "", ""),
                '                .Quantity = 1}
                '            tmpoccurrenceslist.Add(tmpOccItem)
                '        ElseIf occ.SubOccurrences.Count > 0 Then

                '        End If
                '    Next
                'Next

                Dim tmpPartsList = (From ccitem As Document In itemisedPartsList
                                    Where ccitem.FullFileName.Contains("Content Center")
                                    Let foldername As String = IO.Path.GetDirectoryName(ccitem.FullFileName)
                                    Order By foldername Ascending
                                    Select ccitem).ToList()

                ccQuantityList = tmpPartsList.GroupBy(Function(x) x).Where(Function(x) x.Count > 1).Select(Function(x) x.Key).ToList()

                If Not ccPartsList Is Nothing Then
                    If ccBomRowItems Is Nothing Then
                        ccBomRowItems = New List(Of BomRowItem)
                        For Each doc As Document In ccPartsList
                            Dim item As New BomRowItem() With {
                                .ItemNo = ItemNo,
                                .Document = doc.FullFileName,
                                .Material = iProperties.GetorSetStandardiProperty(doc, PropertiesForDesignTrackingPropertiesEnum.kMaterialDesignTrackingProperties, "", ""),
                                .Quantity = 1}
                            'log.Info("Item: " & ItemNo & doc.FullFileName)
                            ccBomRowItems.Add(item)
                            ItemNo += 1
                        Next
                    Else
                        ItemNo = (From m As BomRowItem In ccBomRowItems
                                  Order By m.ItemNo Ascending
                                  Select m.ItemNo).Last()
                        ItemNo += 1
                        For Each doc As Document In ccPartsList
                            Dim testBomRowItem As BomRowItem = (From m As BomRowItem In ccBomRowItems
                                                                Where m.Document = doc.FullFileName
                                                                Select m).FirstOrDefault()
                            If testBomRowItem Is Nothing Then
                                ccBomRowItems.Add(New BomRowItem() With {
                                                  .Document = doc.FullFileName,
                                                  .ItemNo = Convert.ToInt32(ItemNo),
                                                  .Material = iProperties.GetorSetStandardiProperty(doc, PropertiesForDesignTrackingPropertiesEnum.kMaterialDesignTrackingProperties, "", ""),
                                                  .Quantity = 1})
                                ItemNo += 1
                            End If
                        Next
                    End If
                    ConvertBomRowItemsToAttributes()
                    ProcessAllAssemblyOccurrences()
                End If
            End If
        Catch ex As Exception
            log.Error(ex.Message, ex)
        End Try
    End Sub

    Private Function processSubAssys(assyDef As AssemblyComponentDefinition) As List(Of BomRowItem)
        Dim tmplist As List(Of BomRowItem) = New List(Of BomRowItem)
        Try

            'Dim listofThisPart As List(Of Document) = (From thisocc As ComponentOccurrence In assyDef.Occurrences
            '                                           Let thispartDoc As Document = thisocc.Definition.Document
            '                                           Where thispartDoc Is ccpart
            '                                           Select thispartDoc).ToList()

            For Each occ As ComponentOccurrence In assyDef.Occurrences
                If occ.SubOccurrences.Count > 0 Then
                    Dim occDoc As Document = occ.Definition.Document
                    Dim subAssyDef As AssemblyComponentDefinition = occ.Definition
                    tmplist.AddRange(processSubAssys(subAssyDef))
                Else
                    Dim listofPartsInThisAssembly As List(Of Document) = New List(Of Document)
                    For Each ccpart As Document In ccPartsList
                        Dim tmpCompDef As ComponentDefinition = occ.Definition
                        Dim thisDoc As Document = tmpCompDef.Document
                        If TypeOf (thisDoc) Is PartDocument Then
                            If thisDoc Is ccpart Then
                                listofPartsInThisAssembly.Add(thisDoc)
                            End If
                        End If
                    Next
                    Dim groupedbyFilename = listofPartsInThisAssembly.OrderBy(Function(a As Document) a.FullFileName).GroupBy(Function(x As Document) x.FullFileName)
                    For Each partgroup As Object In groupedbyFilename
                        For Each item As Document In partgroup
                            Dim tmpOccItem As New BomRowItem() With {
                                    .Document = item.FullFileName,
                                    .ItemNo = 0,
                                    .Material = iProperties.GetorSetStandardiProperty(item, PropertiesForDesignTrackingPropertiesEnum.kMaterialDesignTrackingProperties, "", ""),
                                    .Quantity = partgroup.Count}
                            tmplist.Add(tmpOccItem)
                        Next
                    Next
                    'If listofThisPart.Count > 0 Then
                    '    Dim tmpOccItem As New BomRowItem() With {
                    '                .Document = ccpart.FullFileName,
                    '                .ItemNo = 0,
                    '                .Material = iProperties.GetorSetStandardiProperty(ccpart, PropertiesForDesignTrackingPropertiesEnum.kMaterialDesignTrackingProperties, "", ""),
                    '                .Quantity = listofThisPart.Count}
                    '    tmplist.Add(tmpOccItem)
                    'End If
                End If
            Next
            'If Not listofThisPart Is Nothing And listofThisPart.Count > 0 Then


            Return tmplist
        Catch ex As Exception
            log.Error(ex.Message, ex)
            Return tmplist
        End Try
    End Function

    Private Sub ConvertBomRowItemsToAttributes()
        Dim standardCCPartAttSet As AttributeSet = Nothing
        For Each item As BomRowItem In ccBomRowItems
            Dim bHasAttSet As Boolean = ThisApplication.ActiveDocument.AttributeSets.NameIsUsed("CCPartNumberSet" & item.ItemNo.ToString())
            If bHasAttSet Then
                'standardCCPartAttSet = ThisApplication.ActiveDocument.AttributeSets.Add("CCPartNumberSet" & item.ItemNo.ToString())
                standardCCPartAttSet = ThisApplication.ActiveDocument.AttributeSets.Item("CCPartNumberSet" & item.ItemNo.ToString())
                'should maybe verify whether the values match what has been captured?
            Else
                standardCCPartAttSet = ThisApplication.ActiveDocument.AttributeSets.Add("CCPartNumberSet" & item.ItemNo.ToString())
                standardCCPartAttSet.Add("FileName", ValueTypeEnum.kStringType, item.Document)
                standardCCPartAttSet.Add("StandardPartNum", ValueTypeEnum.kStringType, item.ItemNo.ToString)
                standardCCPartAttSet.Add("Quantity", ValueTypeEnum.kIntegerType, item.Quantity)
                standardCCPartAttSet.Add("Material", ValueTypeEnum.kStringType, item.Material)
            End If
            'Dim attributenames() As String = {"FileName, StandardPartNum"}
            'Dim valueTypes() As ValueTypeEnum = {ValueTypeEnum.kStringType, ValueTypeEnum.kStringType}
            'Dim attributeValues() As String = {item.Document, item.ItemNo.ToString}
            'Dim standardCCPartAttEnum As AttributesEnumerator = standardCCPartAttSet.AddAttributes(attributenames, valueTypes, attributeValues, False)
            'Dim standardCCPart As AttributeSet = standardCCPartAttSet.AddAttributes(attributenames, valueTypes, attributeValues, False)
        Next
    End Sub

    Public Sub ProcessAllAssemblyOccurrences()
        Try
            Dim oDoc As Inventor.AssemblyDocument = ThisApplication.ActiveDocument
            'Dim AssySubAssemblies As List(Of AssemblyDocument) = Nothing
            Dim oCompDef As Inventor.ComponentDefinition = oDoc.ComponentDefinition
            'process this assembly
            RenumberBomViews(oDoc.ComponentDefinition)
            ' Get all referenced assemblies in one list
            AssySubAssemblies = (From assyDoc As Document In oDoc.AllReferencedDocuments
                                 Where TypeOf (assyDoc) Is AssemblyDocument
                                 Let selectedDoc As AssemblyDocument = assyDoc
                                 Where selectedDoc.ComponentDefinition.BOM.StructuredViewEnabled = True
                                 Select selectedDoc).ToList()
            For Each assy As Document In AssySubAssemblies
                Dim ThisAssy As AssemblyDocument = assy
                RenumberBomViews(ThisAssy.ComponentDefinition)
            Next
        Catch ex As Exception
            log.Error(ex.Message, ex)
        End Try
    End Sub

    Private Sub RenumberBomViews(parentAssyCompDef As AssemblyComponentDefinition)
        Try
            Dim AssyBom As BOM = parentAssyCompDef.BOM
            Dim ParentAssyDoc As AssemblyDocument = parentAssyCompDef.Document
            If Not AssyBom.StructuredViewEnabled = True Then
                log.Info("structured bom view disabled: " & ParentAssyDoc.FullFileName)
            Else
                updatestatusbar("Processing: " & ParentAssyDoc.FullFileName)
                AssyBom.StructuredViewFirstLevelOnly = True
                Dim ThisAssyBOMView As BOMView = AssyBom.BOMViews.Item("Structured")
                RenumberBOMViewRows(ParentAssyDoc, ThisAssyBOMView)
                ThisAssyBOMView = AssyBom.BOMViews.Item("Parts Only")
                RenumberBOMViewRows(ParentAssyDoc, ThisAssyBOMView)
            End If
        Catch ex As Exception
            log.Error(ex.Message, ex)
        End Try
    End Sub

    Public Sub RenumberBOMViewRows(ParentAssyDoc As Document, currentView As BOMView)
        Try
            For Each row As BOMRow In currentView.BOMRows
                Dim RowCompDef As ComponentDefinition = row.ComponentDefinitions(1)
                Dim thisDoc As Document = RowCompDef.Document
                If row.Promoted Then
                    log.Info(thisDoc.FullFileName & " is Promoted!")
                End If
                Dim matchingStoredDocument As BomRowItem = (From m As BomRowItem In ccBomRowItems
                                                            Where m.Document = thisDoc.FullFileName
                                                            Select m).FirstOrDefault()
                If Not matchingStoredDocument Is Nothing Then
                    log.Info("Assembly: " & IO.Path.GetFileName(ParentAssyDoc.FullFileName) &
                             " | Part: " & IO.Path.GetFileName(thisDoc.FullFileName) &
                             " Item No changed from: " & row.ItemNumber &
                             " to: " & matchingStoredDocument.ItemNo)
                    row.ItemNumber = matchingStoredDocument.ItemNo
                    matchingStoredDocument.Quantity = row.TotalQuantity
                Else
                    'assumes we used 6 digits for our COTS numbering!
                    If IO.Path.GetFileNameWithoutExtension(thisDoc.FullFileName).StartsWith("COTS") Then
                        Dim COTSnumber As String = IO.Path.GetFileNameWithoutExtension(thisDoc.FullFileName)
                        Dim COTSBase As Double = 0
                        Dim regex As New Regex("(\d{6})")
                        Dim f As String = String.Empty
                        f = regex.Match(COTSnumber).Captures(0).ToString()
                        COTSBase = GetRoundNum(Convert.ToDouble(f), 100000)
                        Dim COTSNum As Double = Convert.ToDouble(f)
                        Dim newItemNum As String = Convert.ToString(COTSNum - COTSBase)
                        row.ItemNumber = newItemNum
                    Else
                        If Not iProperties.SetorCreateCustomiProperty(thisDoc, "ItemNo") = String.Empty Then
                            row.ItemNumber = iProperties.SetorCreateCustomiProperty(thisDoc, "ItemNo")
                        End If
                    End If

                End If

            Next
            currentView.Sort("Item", True)
        Catch ex As Exception
            log.Error(ex.Message, ex)
        End Try
    End Sub
#End Region

#Region "Insert Dummy Files"
    'Imports Microsoft.office.interop.excel
    Public Sub Main()
        'logging set up
        thisAssemblyPath = System.IO.Path.GetDirectoryName(thisAssembly.Location)
        logHelper.Init()
        logHelper.AddConsoleLogging()
        'the next line works but we want a rolling log.
        logHelper.AddFileLogging(System.IO.Path.Combine(thisAssemblyPath, "GraitecExtensionsServer.log"))
        logHelper.AddFileLogging("C:\Logs\MyLogFile.txt", log4net.Core.Level.All, True)
        logHelper.AddRollingFileLogging("C:\Logs\RollingFileLog.txt", log4net.Core.Level.All, True)
        log.Debug("Loading iLogic External Debug")
        Call insertdummyfiles()
    End Sub

    Public Sub insertdummyfiles()
        'define assembly
        Dim asmDoc As AssemblyDocument
        asmDoc = ThisApplication.ActiveDocument
        'create a transaction to encapsulate all our additions in one undo.
        Dim tr As Transaction
        tr = ThisApplication.TransactionManager.StartTransaction(
            ThisApplication.ActiveDocument,
            "Create Standard Parts From Excel")
        Try
            excelapp = GetOrCreateInstance("Excel.Application")
            'Dim workBook As Workbook = Nothing
            Dim workSheet As Worksheet = Nothing
            Dim ProjectRootFolder As String = ThisApplication.DesignProjectManager.ActiveDesignProject.WorkspacePath
            Dim excelFilename As String = ProjectRootFolder & "\QuantifiableParts.xlsx"
            workBook = excelapp.Workbooks.Open(excelFilename)
            If (workBook IsNot Nothing) Then
                workSheet = workBook.Worksheets(1)
                usedRange = workSheet.UsedRange
            End If
            'workBook.Close()
            'excelapp.Quit()
            'MessageBox.Show(excelfilename)
            Dim COTSPrefix As String = "COTS-"
            Dim COTSPartNumStart As Long = 100000
            Dim ItemNo As Integer = 200
            'commented whilst debugging
            Dim folderbrowser As New System.Windows.Forms.FolderBrowserDialog()
            folderbrowser.RootFolder = System.Environment.SpecialFolder.MyComputer
            folderbrowser.Description = "Select Folder to look for files to process."
            folderbrowser.ShowDialog()
            Dim SelectedProjectFolder As String = folderbrowser.SelectedPath
            'Dim SelectedProjectFolder As String = "C:\Users\Alex.Fielder\OneDrive\Inventor\Designs\Test"
            If SelectedProjectFolder Is Nothing Then Exit Sub
            Dim COTSInitialPrefix As String = InputBox("What number do you want to start at?", "Title", CStr(COTSPartNumStart))
            COTSPartNumStart = Convert.ToInt32(COTSInitialPrefix)
            'GoExcel.Open(excelFilename, "Sheet1")
            'get iProperties from the XLS file
            For MyRow = 3 To usedRange.Rows.Count 'index row 3 through 1000

                If Not TypeOf (usedRange.Cells(MyRow, 5).Value2) Is Double Then
                    If usedRange.Cells(MyRow, 5).Value2 = "" Then
                        'If usedRange.Cells(MyRow, 5).Value2 = "" Then
                        ItemNo += 100
                        ItemNo = GetRoundNum(ItemNo, 100)
                        If ItemNo = 500 Then 'Content Centre
                            ItemNo += 100
                        End If
                        'reset the main counter
                        COTSPartNumStart = Convert.ToInt32(COTSInitialPrefix)
                        Continue For
                    End If
                End If
                Dim PartNum As String = COTSPrefix & (COTSPartNumStart + ItemNo)
                'GoExcel.CellValue("B" & MyRow)	'PART NUMBER
                Dim Quantity As Double
                If Not usedRange.Cells(MyRow, 3).Value2 = 1 Then
                    Quantity = usedRange.Cells(MyRow, 3).Value2   'UNIT QUANTITY	

                Else
                    Quantity = 1
                End If
                Dim Description As String = usedRange.Cells(MyRow, 6).Value2 & " - " & usedRange.Cells(MyRow, 8).Value2 'DESIGNATION & " - " & DESCRIPTION
                'Dim iProp6 as String = GoExcel.CellValue("F" & MyRow)	'VENDOR
                'Dim iProp7 as String = GoExcel.CellValue("G" & MyRow)	'REV
                'Dim iProp8 as String = GoExcel.CellValue("H" & MyRow)	'COMMENTS
                'Dim ItemNo As String = GoExcel.CellValue("I" & MyRow)	'ITEM NUMBER
                'Dim iProp10 as String = GoExcel.CellValue("K" & MyRow)	'SUBJECT/LEGACY DRAWING NUMBER
                Dim occs As ComponentOccurrences
                occs = asmDoc.ComponentDefinition.Occurrences
                'sets up a Matrix based on the origin of the Assembly - we could translate each insert away from 0,0,0 but there's no real need to do that.
                Dim PosnMatrix As Matrix
                PosnMatrix = ThisApplication.TransientGeometry.CreateMatrix
                Dim basefilename = ProjectRootFolder & "\Graitec\DT-PINK_DISC-000.ipt"
                Dim newfilename As String = SelectedProjectFolder & "\" & PartNum & ".ipt"
                '		MessageBox.Show(newfilename, "Title") 'for debuggering!
                '		Exit Sub
                If Not System.IO.File.Exists(newfilename) Then 'we need to create it
                    updatestatusbar("Creating " & newfilename)
                    System.IO.File.Copy(basefilename, newfilename)
                End If
                'creates a componentoccurence object
                Dim realOcc As ComponentOccurrence
                'and adds it at the origin of the assembly.
                realOcc = occs.Add(newfilename, PosnMatrix)
                Dim docToEditiProps = realOcc.Definition.Document
                Dim realOccStr As String = realOcc.Name
                'Assign iProperties
                iProperties.GetorSetStandardiProperty(
                    docToEditiProps,
                    PropertiesForDesignTrackingPropertiesEnum.kDescriptionDesignTrackingProperties,
                    Description)
                'iProperties.Value(realOccStr, "Project", "Description") = Description
                'iProperties.Value(realOccStr, "Project", "Part Number") = PartNum
                iProperties.GetorSetStandardiProperty(
                    docToEditiProps,
                    PropertiesForDesignTrackingPropertiesEnum.kPartNumberDesignTrackingProperties,
                    PartNum)
                iProperties.GetorSetStandardiProperty(
                    docToEditiProps,
                    PropertiesForSummaryInformationEnum.kRevisionSummaryInformation,
                    "A")
                iProperties.SetorCreateCustomiProperty(docToEditiProps, "ItemNo", ItemNo)
                'iProperties.Value(realOccStr, "Project", "Revision Number") = "A"
                'End Assign iProperties
                realOcc.Visible = False 'hide the first instance
                Dim index As Integer
                index = 2
                Do While index <= CInt(Quantity)
                    Dim tmpOcc As ComponentOccurrence
                    tmpOcc = occs.AddByComponentDefinition(realOcc.Definition, PosnMatrix)
                    tmpOcc.Visible = False ' and all subsequent occurrences.
                    index += 1
                Loop
                COTSPartNumStart += 1
            Next
        Catch ex As Exception
            log.Error(ex.Message, ex)
            workBook.Close()
            excelapp.Quit()
            MessageBox.Show("Error: " & ex.Message)
            tr.Abort()
            m_DocToUpdate.Document.DocumentUpdate()
        Finally
            If excelapp IsNot Nothing Then
                excelapp = Nothing
            End If
            If tr IsNot Nothing Then
                tr.End()
            End If
            m_DocToUpdate.Document.DocumentUpdate()
        End Try
    End Sub

    Private Function GetOrCreateInstance(appName As String) As Excel.Application
        Try
            Return GetInstance(appName)
        Catch ex As Exception
            Return CreateInstance(appName)
        End Try
    End Function

#End Region
    Private Function CreateInstance(appName As String) As Excel.Application
        Return Activator.CreateInstance(Type.GetTypeFromProgID(appName))
    End Function

    Private Function GetInstance(appName As String) As Excel.Application
        Return Marshal.GetActiveObject(appName)
    End Function

    Private Function GetRoundNum(ByVal Number As Double, ByVal multiple As Integer) As Double
        GetRoundNum = CInt(Number / multiple) * multiple
    End Function
    Sub updatestatusbar(ByVal message As String)
        ThisApplication.StatusBarText = message
    End Sub
    Sub updatestatusbar(ByVal percent As Double, ByVal message As String)
        ThisApplication.StatusBarText = message + " (" & percent.ToString("P1") + ")"
    End Sub
End Class

#Region "Helper classes"

Public Class viewstate
    Public viewEnabled As Boolean
    Public view As DrawingView
    'Public curves as DrawingCurvesEnumerator
End Class

Public Class iProperties

    Private thisAssembly As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
    Private thisAssemblyPath As String = String.Empty
    Private logHelper As Log4NetFileHelper = New Log4NetFileHelper()
    Private Shared ReadOnly log As ILog = LogManager.GetLogger(GetType(iProperties))

    Public Shared Function GetiPropertyDisplayName(ByVal iProp As Inventor.Property) As String
        Return iProp.DisplayName
    End Function

    Public Shared Function GetiPropertyType(ByVal iProp As Inventor.Property) As ObjectTypeEnum
        Return iProp.Type
    End Function

    Public Shared Function GetiPropertyTypeString(ByVal iProp As Inventor.Property) As String
        Dim valToTest As Object = iProp.Value
        Dim intResult As Integer = Nothing
        If Integer.TryParse(iProp.Value, intResult) Then
            Return "Number"
        End If

        Dim doubleResult As Double = Nothing
        If Double.TryParse(iProp.Value, doubleResult) Then
            Return "Number"
        End If

        Dim dateResult As Date = Nothing
        If Date.TryParse(iProp.Value, dateResult) Then
            Return "Date"
        End If

        Dim booleanResult As Boolean = Nothing
        If Boolean.TryParse(iProp.Value, booleanResult) Then
            Return "Boolean"
        End If

        'Dim currencyResult As Currency = Nothing

        'should probably do this last as most property values will equate to string!
        Dim strResult As String = String.Empty
        If Not iProp.Value.ToString() = String.Empty Then
            Return "String"
        End If
        Return Nothing
    End Function
#Region "Set iProperty Values"
#Region "Get or Set Standard iProperty Values"

    ''' <summary>
    ''' Design Tracking Properties
    ''' </summary>
    ''' <param name="DocToUpdate"></param>
    ''' <param name="iPropertyTypeEnum"></param>
    ''' <param name="newpropertyvalue"></param>
    ''' <returns></returns>
    Public Shared Function GetorSetStandardiProperty(ByVal DocToUpdate As Inventor.Document,
                                                    ByVal iPropertyTypeEnum As PropertiesForDesignTrackingPropertiesEnum,
                                                    Optional ByRef newpropertyvalue As String = "",
                                                    Optional ByRef propertyTypeStr As String = "") As String
        Dim invProjProperties As PropertySet = DocToUpdate.PropertySets.Item("{32853F0F-3444-11D1-9E93-0060B03C1CA6}")
        Dim currentvalue As String = String.Empty
        If Not newpropertyvalue = String.Empty Then
            invProjProperties.ItemByPropId(iPropertyTypeEnum).Value = newpropertyvalue.ToString()
        Else
            currentvalue = invProjProperties.ItemByPropId(iPropertyTypeEnum).Value
            newpropertyvalue = GetiPropertyDisplayName(invProjProperties.ItemByPropId(iPropertyTypeEnum))
        End If
        If propertyTypeStr = String.Empty Then
            propertyTypeStr = GetiPropertyTypeString(invProjProperties.ItemByPropId(iPropertyTypeEnum))
        End If
        Return currentvalue
    End Function

    ''' <summary>
    ''' Document Summary Properties
    ''' </summary>
    ''' <param name="DocToUpdate"></param>
    ''' <param name="iPropertyTypeEnum"></param>
    ''' <param name="newpropertyvalue"></param>
    ''' <returns></returns>
    Public Shared Function GetorSetStandardiProperty(ByVal DocToUpdate As Inventor.Document,
                                                    ByVal iPropertyTypeEnum As PropertiesForDocSummaryInformationEnum,
                                                    Optional ByRef newpropertyvalue As String = "",
                                                    Optional ByRef propertyTypeStr As String = "") As String
        Dim invDocSummaryProperties As PropertySet = DocToUpdate.PropertySets.Item("{D5CDD502-2E9C-101B-9397-08002B2CF9AE}")
        Dim currentvalue As String = String.Empty
        If Not newpropertyvalue = String.Empty Then
            invDocSummaryProperties.ItemByPropId(iPropertyTypeEnum).Value = newpropertyvalue.ToString()
        Else
            currentvalue = invDocSummaryProperties.ItemByPropId(iPropertyTypeEnum).Value
            newpropertyvalue = GetiPropertyDisplayName(invDocSummaryProperties.ItemByPropId(iPropertyTypeEnum))
        End If
        If propertyTypeStr = String.Empty Then
            propertyTypeStr = GetiPropertyTypeString(invDocSummaryProperties.ItemByPropId(iPropertyTypeEnum))
        End If
        Return currentvalue
    End Function

    ''' <summary>
    ''' Summary Properties
    ''' </summary>
    ''' <param name="DocToUpdate"></param>
    ''' <param name="iPropertyTypeEnum"></param>
    ''' <param name="newpropertyvalue"></param>
    ''' <returns></returns>
    Public Shared Function GetorSetStandardiProperty(ByVal DocToUpdate As Inventor.Document,
                                                    ByVal iPropertyTypeEnum As PropertiesForSummaryInformationEnum,
                                                    Optional ByRef newpropertyvalue As String = "",
                                                    Optional ByRef propertyTypeStr As String = "") As String
        Dim invSummaryiProperties As PropertySet = DocToUpdate.PropertySets.Item("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}")
        Dim currentvalue As String = String.Empty
        If Not newpropertyvalue = String.Empty Then
            invSummaryiProperties.ItemByPropId(iPropertyTypeEnum).Value = newpropertyvalue.ToString()
        Else
            currentvalue = invSummaryiProperties.ItemByPropId(iPropertyTypeEnum).Value
            newpropertyvalue = GetiPropertyDisplayName(invSummaryiProperties.ItemByPropId(iPropertyTypeEnum))
        End If
        If propertyTypeStr = String.Empty Then
            propertyTypeStr = GetiPropertyTypeString(invSummaryiProperties.ItemByPropId(iPropertyTypeEnum))
        End If
        Return currentvalue
    End Function
#End Region
#Region "Get or Set Custom iProperty Values"

    ''' <summary>
    ''' This method should set or get any custom iProperty value
    ''' </summary>
    ''' <param name="Doc">the document to edit</param>
    ''' <param name="PropertyName">the iProperty name to retrieve or update</param>
    ''' <param name="PropertyValue">the optional value to assign - if empty we are retrieving a value</param>
    ''' <returns></returns>
    Friend Shared Function SetorCreateCustomiProperty(ByVal Doc As Inventor.Document, ByVal PropertyName As String, Optional ByVal PropertyValue As Object = Nothing) As Object
        ' Get the custom property set.
        Dim customPropSet As Inventor.PropertySet
        Dim customproperty As Object = Nothing
        Try
            customPropSet = Doc.PropertySets.Item("Inventor User Defined Properties")

            ' Get the existing property, if it exists.
            Dim prop As Inventor.Property = Nothing
            Dim propExists As Boolean = True
            Try
                prop = customPropSet.Item(PropertyName)
            Catch ex As Exception
                propExists = False
            End Try
            If Not PropertyValue Is Nothing Then
                ' Check to see if the property was successfully obtained.
                If Not propExists Then
                    ' Failed to get the existing property so create a new one.
                    If Not Doc.FullFileName.Contains("Content Center") Then
                        prop = customPropSet.Add(PropertyValue, PropertyName)
                    End If
                Else
                    If Not prop Is Nothing Then ' avoids the error where the item *should* have the item number but perhaps doesn't because reasons!
                        ' Change the value of the existing property.
                        prop.Value = PropertyValue
                    End If
                End If
            Else
                customproperty = prop.Value
            End If
        Catch ex As Exception
            Log.Error(ex.Message, ex)
        End Try
        Return customproperty

    End Function
#End Region

#End Region
End Class

''' <summary>
''' All of these functions will work with the currently active document inside of Inventor.
''' </summary>
Public Class Parameters
    'Public Shared ThisApplication As Inventor.Application
#Region "Parameters"
    ''' <summary>
    ''' UNTESTED 2016-05-24 AF
    ''' Sets a string parameter value
    ''' </summary>
    ''' <param name="ParameterName"></param>
    ''' <param name="ParameterValue"></param>
    Public Shared Sub SetParameter(ByVal Doc As Inventor.Document, ByVal ParameterName As String, ByVal ParameterValue As String)
        ' Get the Parameters object. Assumes a part or assembly document is active.
        Dim oParameters As Inventor.Parameters = Doc.ComponentDefinition.Parameters

        ' Get the parameter named "Length".
        Dim oLengthParam As Inventor.Parameter = oParameters.Item(ParameterName)

        ' Change the equation of the parameter.
        oLengthParam.Expression = ParameterValue

        ' Update the document.
        'If Power_Pack_For_Inventor_AddIn.UpdateAfterEachParameterChange Then
        Doc.Update()
        'End If
    End Sub
    ''' <summary>
    ''' UNTESTED 2016-05-24 AF
    ''' Sets a number parameter value
    ''' </summary>
    ''' <param name="ParameterName"></param>
    ''' <param name="ParameterValue"></param>
    Public Shared Sub SetParameter(ByVal Doc As Inventor.Document, ByVal ParameterName As String, ByVal ParameterValue As Double)
        ' Get the Parameters object. Assumes a part or assembly document is active.
        Dim oParameters As Inventor.Parameters = Doc.ComponentDefinition.Parameters

        ' Get the parameter named "Length".
        Dim oLengthParam As Inventor.Parameter = oParameters.Item(ParameterName)

        ' Change the equation of the parameter.
        oLengthParam.Expression = ParameterValue

        ' Update the document.
        'If Power_Pack_For_Inventor_AddIn.UpdateAfterEachParameterChange Then
        Doc.Update()
        'End If
    End Sub
    ''' <summary>
    ''' UNTESTED 2016-05-24 AF
    ''' Sets a true/false parameter value
    ''' </summary>
    ''' <param name="ParameterName"></param>
    ''' <param name="ParameterValue"></param>
    Public Shared Sub SetParameter(ByVal Doc As Inventor.Document, ByVal ParameterName As String, ByVal ParameterValue As Boolean)
        ' Get the Parameters object. Assumes a part or assembly document is active.
        Dim oParameters As Inventor.Parameters = Doc.ComponentDefinition.Parameters

        ' Get the parameter named "Length".
        Dim oLengthParam As Inventor.Parameter = oParameters.Item(ParameterName)

        ' Change the equation of the parameter.
        oLengthParam.Expression = ParameterValue

        ' Update the document.
        'If Power_Pack_For_Inventor_AddIn.UpdateAfterEachParameterChange Then
        Doc.Update()
        'End If
    End Sub
    ''' <summary>
    ''' UNTESTED 2016-05-24 AF
    ''' Sets a Date Parameter Value
    ''' </summary>
    ''' <param name="ParameterName"></param>
    ''' <param name="ParameterValue"></param>
    Public Shared Sub SetParameter(ByVal Doc As Inventor.Document, ByVal ParameterName As String, ByVal ParameterValue As DateTime)
        ' Get the Parameters object. Assumes a part or assembly document is active.
        Dim oParameters As Inventor.Parameters = Doc.ComponentDefinition.Parameters

        ' Get the parameter named "Length".
        Dim oLengthParam As Inventor.Parameter = oParameters.Item(ParameterName)

        ' Change the equation of the parameter.
        oLengthParam.Expression = ParameterValue

        ' Update the document.
        'If Power_Pack_For_Inventor_AddIn.UpdateAfterEachParameterChange Then
        Doc.Update()
        'End If
    End Sub

    Public Shared Function GetParameter(ByVal Doc As Inventor.Document, ByVal ParamName As String) As Inventor.Parameter
        Dim oParameters As Inventor.Parameters = Doc.ComponentDefinition.Parameters
        If oParameters(ParamName).ParameterType = ParameterTypeEnum.kUserParameter Then
            Return GetUserParameter(Doc, ParamName)
        ElseIf oParameters(ParamName).ParameterType = ParameterTypeEnum.kReferenceParameter Then
            Return GetReferenceParameter(Doc, ParamName)
        ElseIf oParameters(ParamName).ParameterType = ParameterTypeEnum.kModelParameter Then
            Return GetModelParameter(Doc, ParamName)
        ElseIf oParameters(ParamName).ParameterType = ParameterTypeEnum.kDerivedParameter Then
            Return GetDerivedParameter(Doc, ParamName)
        ElseIf oParameters(ParamName).ParameterType = ParameterTypeEnum.kTableParameter Then
            Return GetTableParameter(Doc, ParamName)
            'Throw New NotSupportedException()
        End If
        Return Nothing
    End Function

    ''' <summary>
    ''' Gets the object of a parameter by name
    ''' </summary>
    ''' <param name="ParameterName"></param>
    ''' <returns></returns>
    Public Shared Function GetUserParameter(ByVal Doc As Inventor.Document, ByVal ParameterName As String) As Inventor.UserParameter
        Dim oParameters As Inventor.Parameters = Doc.ComponentDefinition.Parameters
        Return oParameters.Item(ParameterName)
    End Function

    ''' <summary>
    ''' Gets the object of a reference parameter by name
    ''' </summary>
    ''' <param name="ParameterName"></param>
    ''' <returns></returns>
    Public Shared Function GetReferenceParameter(ByVal Doc As Inventor.Document, ByVal ParameterName As String) As Inventor.ReferenceParameter
        Dim oParameters As Inventor.Parameters = Doc.ComponentDefinition.Parameters
        Return oParameters.Item(ParameterName)
    End Function

    ''' <summary>
    ''' Gets the object of a model parameter by name
    ''' </summary>
    ''' <param name="ParameterName"></param>
    ''' <returns></returns>
    Public Shared Function GetModelParameter(ByVal Doc As Inventor.Document, ByVal ParameterName As String) As Inventor.ModelParameter
        Dim oParameters As Inventor.Parameters = Doc.ComponentDefinition.Parameters
        Return oParameters.Item(ParameterName)
    End Function

    ''' <summary>
    ''' Gets the object of a derived parameter by name
    ''' </summary>
    ''' <param name="ParameterName"></param>
    ''' <returns></returns>
    Public Shared Function GetDerivedParameter(ByVal Doc As Inventor.Document, ByVal ParameterName As String) As Inventor.DerivedParameter
        Dim oParameters As Inventor.Parameters = Doc.ComponentDefinition.Parameters
        Return oParameters.Item(ParameterName)
    End Function

    ''' <summary>
    ''' Gets the object of a Table Parameter NOT IMPLEMENTED AS OF 2016-09-28!
    ''' </summary>
    ''' <param name="ParameterName"></param>
    ''' <returns></returns>
    Public Shared Function GetTableParameter(ByVal Doc As Inventor.Document, ByVal ParameterName As String) As Inventor.TableParameter
        Throw New NotImplementedException
    End Function

#End Region

End Class
#End Region
