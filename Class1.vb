Option Explicit On
Imports System.IO
Imports System.Windows.Forms
Imports Inventor
Imports Autodesk.iLogic.Interfaces
Imports iLogicExternalDebug
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices
Imports log4net

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

#Region "Renumber Item lists"
    Public Sub BeginRenumberItems()
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

    Private ccPartsList As List(Of Document) = Nothing
    Private ccBomRowItems As List(Of BomRowItem) = Nothing
    Private ItemNo As Integer = 500
    Private Sub RunRenumberItems()

        If TypeOf (ThisApplication.ActiveDocument) Is AssemblyDocument Then
            Dim AssyDoc As AssemblyDocument = ThisApplication.ActiveDocument
            ccPartsList = (From ccDoc As Document In AssyDoc.AllReferencedDocuments
                           Where ccDoc.FullFileName.Contains("Content Center")
                           Let foldername As String = IO.Path.GetDirectoryName(ccDoc.FullFileName)
                           Order By foldername Ascending
                           Select ccDoc).Distinct().ToList()
            ccPartsList.RemoveAll(Function(x As PartDocument) x.ComponentDefinition.BOMStructure = BOMStructureEnum.kReferenceBOMStructure)
            If Not ccPartsList Is Nothing Then
                If ccBomRowItems Is Nothing Then
                    ccBomRowItems = New List(Of BomRowItem)
                    For Each doc As Document In ccPartsList
                        Dim item As New BomRowItem() With {
                            .ItemNo = ItemNo,
                            .Document = doc.FullFileName}
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
                                              .ItemNo = Convert.ToInt32(ItemNo)})
                            ItemNo += 1
                        End If
                    Next
                End If
                ConvertBomRowItemsToAttributes()
                ProcessAllAssemblyOccurrences()
            End If
        End If
    End Sub

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
            Dim AssyPartOccurrences As List(Of ComponentOccurrence) = Nothing
            Dim oCompDef As Inventor.ComponentDefinition = oDoc.ComponentDefinition
            Dim sMsg As String
            Dim iLeafNodes As Long
            Dim iSubAssemblies As Long
            ' Get all occurrences from component definition for Assembly document
            'AssyPartOccurrences = (From partDoc As ComponentOccurrence In oCompDef.Occurrences
            '                       Where TypeOf (partDoc.Definition.Document) Is PartDocument
            '                       Select partDoc).ToList()


            For Each oCompOcc As ComponentOccurrence In oCompDef.Occurrences
                If oCompOcc.SubOccurrences.Count = 0 Then
                    iLeafNodes = iLeafNodes + 1
                    RenumberBomOccurrences(oCompDef, oCompOcc)
                Else
                    iSubAssemblies = iSubAssemblies + 1
                    Call processAllSubOcc(oCompOcc,
                                    sMsg,
                                    iLeafNodes,
                                    iSubAssemblies)
                End If
            Next
        Catch ex As Exception
            log.Error(ex.Message, ex)
        End Try
    End Sub

    ' This function is called for processing sub assembly.  It is called recursively
    ' to iterate through the entire assembly tree.
    Private Sub processAllSubOcc(ByVal oCompOcc As ComponentOccurrence,
                                 ByRef sMsg As String,
                                 ByRef iLeafNodes As Long,
                                 ByRef iSubAssemblies As Long)
        Try
            For Each oSubCompOcc As ComponentOccurrence In oCompOcc.SubOccurrences
                If oSubCompOcc.BOMStructure = BOMStructureEnum.kReferenceBOMStructure Then
                    Continue For
                End If
                ' Check if it's child occurrence (leaf node)
                If oSubCompOcc.SubOccurrences.Count = 0 Then
                    'Debug.Print oSubCompOcc.Name
                    RenumberBomOccurrences(oCompOcc.Definition, oSubCompOcc)
                Else
                    sMsg = sMsg + oSubCompOcc.Name + vbCr
                    iSubAssemblies = iSubAssemblies + 1

                    Call processAllSubOcc(oSubCompOcc,
                                          sMsg,
                                          iLeafNodes,
                                          iSubAssemblies)
                End If
            Next
        Catch ex As Exception

        End Try
    End Sub

    Private Sub RenumberBomOccurrences(parentAssyCompDef As AssemblyComponentDefinition, partOccurrence As ComponentOccurrence)
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
                Else
                    If Not iProperties.SetorCreateCustomiProperty(thisDoc, "ItemNo") = String.Empty Then
                        row.ItemNumber = iProperties.SetorCreateCustomiProperty(thisDoc, "ItemNo")
                    End If
                End If

            Next
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
