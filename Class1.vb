Imports System.Windows.Forms
Imports Inventor
Imports Autodesk.iLogic.Interfaces
Imports iLogicExternalDebug

Public Class ExtClass
#Region "Properties"
    ''' <summary>
    ''' useful for passing the inventor application to this .dll
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared m_inventorApplication As Inventor.Application

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
    Public Sub Update()


        NumFiltersWide = Parameters.GetParameter(DocToUpdate.Document, "NumFiltersWide")
        NumFiltersHigh = Parameters.GetParameter(DocToUpdate.Document, "NumFiltersHigh")
        NumFiltrationStages = Parameters.GetParameter(DocToUpdate.Document, "FilterHouseNumFiltrationStages")
        NumElementsWide = Parameters.GetParameter(DocToUpdate.Document, "NumElementsWide")
        Dim ThisModule As FilterModule = New FilterModule(
        NumFiltersWide.Value.ToString(),
        NumFiltersHigh.Value.ToString(),
        NumFiltrationStages.Value.ToString(),
        NumElementsWide.Value.ToString())
        ThisModule.NumFiltersThisModule = Convert.ToInt32(NumFiltersWide.Value.ToString.Remove(" ul")) *
            Convert.ToInt32(NumFiltersHigh.Value.ToString.Remove(" ul")) *
            Convert.ToInt32(NumFiltrationStages.Value.ToString.Remove(" ul"))

        Select Case iProperties.SetorCreateCustomiProperty(DocToUpdate.Document, "BasePartNumber")
            Case "FilterModuleMaster" ' our master part file so this better not break anything and updating it first makes sense!
                'UpdateFilterModuleMaster()
                UpdateFilterModuleMaster(ThisModule)
            Case "Master-FrontUpperFlange"
                UpdateFrontUpperFlange("Master-FrontUpperFlange")
            Case "Master-FrontBottomFlange"
                UpdateFrontBottomFlange("Master-FrontBottomFlange")
            Case "Master-RearBottomFlange"
                UpdateRearBottomFlange("Master-RearBottomFlange")
            Case "Master-InterBottomFlange"
                UpdateInterBottomFlange("Master-InterBottomFlange")
            Case "Master-RearUpperFlange"
                UpdateRearUpperFlange("Master-RearUpperFlange")
            Case "Master-InterUpperFlange"
                UpdateInterUpperFlange("Master-InterUpperFlange")
            Case "Master-Stage1LeftVertical"
                UpdateStage1LeftVertical("Master-Stage1LeftVertical")
            Case "Add more as required"

            Case Else

        End Select


    End Sub



    ''' <summary>
    ''' not sure about this method yet as it relies on a lot of chicken !± egg information.
    ''' </summary>
    ''' <param name="thisModule"></param>
    Private Sub UpdateFilterModuleMaster(thisModule As FilterModule)
        If TypeOf DocToUpdate.Document Is PartDocument Then
            'Dim derivedpartcheck As PartDocument = DocToUpdate.Document
            'If derivedpartcheck.ComponentDefinition.ReferenceComponents.DerivedPartComponents.Count = 0 Then
            Dim Element1NumFilters As Parameter = Parameters.GetParameter(DocToUpdate.Document, "MasterElement1NumFilters")
            Dim Element2NumFilters As Parameter = Parameters.GetParameter(DocToUpdate.Document, "MasterElement2NumFilters")
            Dim Element3NumFilters As Parameter = Parameters.GetParameter(DocToUpdate.Document, "MasterElement3NumFilters")
            Dim Element4NumFilters As Parameter = Parameters.GetParameter(DocToUpdate.Document, "MasterElement4NumFilters")
            If Not thisModule.ElementFour Is Nothing Then
                Element4NumFilters.Value = thisModule.ElementFour.NumFilters
            End If
            If Not thisModule.ElementThree Is Nothing Then
                Element3NumFilters.Value = thisModule.ElementThree.NumFilters
            End If

            Element1NumFilters.Value = thisModule.ElementOne.NumFilters
            Element2NumFilters.Value = thisModule.ElementTwo.NumFilters
            Parameters.SetParameter(DocToUpdate.Document, "NumElementsWide", thisModule.NumElementsWide)

            Select Case NumFiltrationStages.Value
                Case 2

                Case 3

                Case 4

                Case Else

            End Select

            'Else
            '    'change something in one of the other files that isn't derived from the master part file.


            'End If
        End If
    End Sub
    'Key Parameters
    ''' <summary>
    ''' The width of this module in Filters
    ''' </summary>
    Public NumFiltersWide As Parameter = Nothing
    ''' <summary>
    ''' The height of this module in Filters
    ''' </summary>
    Public NumFiltersHigh As Parameter = Nothing
    ''' <summary>
    ''' The number of filtration stages in this and other modules for this project.
    ''' </summary>
    Public NumFiltrationStages As Parameter = Nothing
    ''' <summary>
    ''' The number of Elements in this module.
    ''' </summary>
    Public NumElementsWide As Parameter = Nothing

    Sub UpdateFilterModuleMaster()
        'MessageBox.Show("Hello World")
        UpdateKeyWidthParameters()

    End Sub



    ''' <summary>
    ''' Sorts out the standard options for each width-based member
    ''' </summary>
    ''' <param name="MasterpartName">The "MASTER" name of the part we are editing</param>
    ''' <param name="ElementNum">The Element Number to edit.</param>
    ''' <param name="NumHoles">The Number of Holes in this element pattern.</param>
    ''' <param name="HoleOffset">The Offset distance towards the left of the member from the Section intersection.</param>
    ''' <param name="PatternNumSlots">The number of slots in this element pattern.</param>
    ''' <param name="PatternStart">The pattern start value for either the holes or slots.</param>
    Sub UpdateKeyParameters(ByVal MasterpartName As String,
                            ByVal ElementNum As Integer,
                            ByVal NumHoles As Integer,
                            ByVal HoleOffset As String,
                            ByVal PatternNumSlots As Integer,
                            ByVal PatternStart As String)

        'default options
        If MasterpartName.Equals("Master-FrontUpperFlange") Then
            Parameters.SetParameter(DocToUpdate.Document, "Element" & ElementNum.ToString() & "NumHoles", NumHoles.ToString() & " + 1 ul")
            Parameters.SetParameter(DocToUpdate.Document, "Element" & ElementNum.ToString() & "HoleOffset", HoleOffset)
            Parameters.SetParameter(DocToUpdate.Document, "Element" & ElementNum.ToString() & "PatternNumSlots", PatternNumSlots.ToString())
            Parameters.SetParameter(DocToUpdate.Document, "Element" & ElementNum.ToString() & "PatternStart", PatternStart)
            'ElseIf MasterpartName.Equals("Master-FrontBottomFlange") Then
            '    Parameters.SetParameter(DocToUpdate.Document,
            '                            "Element" & ElementNum.ToString() & "Pattern" & PatternNum.ToString() & "Start", PatternStart)
            '    Parameters.SetParameter(DocToUpdate.Document,
            '                            "Element" & ElementNum.ToString() & "Pattern" & PatternNum.ToString() & "NumSlots", PatternNumSlots)
        ElseIf MasterpartName.Equals("Master-FrontBottomFlange") Then
            Parameters.SetParameter(DocToUpdate.Document, "Element" & ElementNum.ToString() & "NumHoles", NumHoles.ToString() & " + 1 ul")
            Parameters.SetParameter(DocToUpdate.Document, "Element" & ElementNum.ToString() & "HoleOffset", HoleOffset)
            Parameters.SetParameter(DocToUpdate.Document, "Element" & ElementNum.ToString() & "PatternNumSlots", PatternNumSlots.ToString())
            Parameters.SetParameter(DocToUpdate.Document, "Element" & ElementNum.ToString() & "PatternStart", PatternStart)
        Else

        End If


    End Sub

    ''' <summary>
    ''' Takes a bunch of inputs and updates parameters
    ''' </summary>
    ''' <param name="MasterpartName"></param>
    ''' <param name="ElementNum"></param>
    ''' <param name="PatternNum"></param>
    ''' <param name="PatternNumSlots"></param>
    ''' <param name="PatternStart"></param>
    ''' <param name="PatternSpacing"></param>
    Sub UpdateKeyPatternParameters(ByVal MasterpartName As String,
                                   ByVal ElementNum As Integer,
                                   ByVal PatternNum As Integer,
                                   ByVal PatternNumSlots As Integer,
                                   ByVal PatternStart As String,
                                   ByVal PatternSpacing As String)
        Parameters.SetParameter(DocToUpdate.Document, "Element" & ElementNum.ToString() & "Pattern" & PatternNum.ToString() & "Start", PatternStart)
        Parameters.SetParameter(DocToUpdate.Document, "Element" & ElementNum.ToString() & "Pattern" & PatternNum.ToString() & "NumSlots", PatternNumSlots)
        Parameters.SetParameter(DocToUpdate.Document, "Element" & ElementNum.ToString() & "Pattern" & PatternNum.ToString() & "Spacing", PatternSpacing)

    End Sub

#Region "Horizontal Members"
    ''' <summary>
    ''' Based in part on this post: http://forums.autodesk.com/t5/inventor-customization/check-if-it-is-derived-part-in-vba/td-p/5147080
    ''' </summary>
    Private Sub UpdateKeyWidthParameters()
        If TypeOf DocToUpdate.Document Is PartDocument Then
            Dim derivedpartcheck As PartDocument = DocToUpdate.Document
            If derivedpartcheck.ComponentDefinition.ReferenceComponents.DerivedPartComponents.Count = 0 Then
                Dim Element1NumFilters As Parameter = Parameters.GetParameter(DocToUpdate.Document, "MasterElement1NumFilters")
                Dim Element2NumFilters As Parameter = Parameters.GetParameter(DocToUpdate.Document, "MasterElement2NumFilters")
                Dim Element3NumFilters As Parameter = Parameters.GetParameter(DocToUpdate.Document, "MasterElement3NumFilters")
                Dim Element4NumFilters As Parameter = Parameters.GetParameter(DocToUpdate.Document, "MasterElement4NumFilters")
                Select Case NumFiltersWide.Value
                    Case 7
                        Parameters.SetParameter(DocToUpdate.Document, "NumElementsWide", "2 ul")
                        Element1NumFilters.Value = 4
                        Element2NumFilters.Value = 3
                    Case 8
                        Parameters.SetParameter(DocToUpdate.Document, "NumElementsWide", "2 ul")
                        Element1NumFilters.Value = 4
                        Element2NumFilters.Value = 4
                    Case 9
                        Parameters.SetParameter(DocToUpdate.Document, "NumElementsWide", "3 ul")
                        Element1NumFilters.Value = 3
                        Element2NumFilters.Value = 3
                        Element3NumFilters.Value = 3
                    Case 10
                        Parameters.SetParameter(DocToUpdate.Document, "NumElementsWide", "3 ul")
                        Element1NumFilters.Value = 3
                        Element2NumFilters.Value = 4
                        Element3NumFilters.Value = 3
                    Case 11
                        Parameters.SetParameter(DocToUpdate.Document, "NumElementsWide", "3 ul")
                        Element1NumFilters.Value = 4
                        Element2NumFilters.Value = 3
                        Element3NumFilters.Value = 4
                    Case 12
                        Parameters.SetParameter(DocToUpdate.Document, "NumElementsWide", "3 ul")
                        Element1NumFilters.Value = 4
                        Element2NumFilters.Value = 4
                        Element3NumFilters.Value = 4
                    Case 13
                        Parameters.SetParameter(DocToUpdate.Document, "NumElementsWide", "4 ul")
                        Element1NumFilters.Value = 3
                        Element2NumFilters.Value = 3
                        Element3NumFilters.Value = 4
                        Element4NumFilters.Value = 3
                    Case 14
                        Parameters.SetParameter(DocToUpdate.Document, "NumElementsWide", "4 ul")
                        Element1NumFilters.Value = 3
                        Element2NumFilters.Value = 4
                        Element3NumFilters.Value = 4
                        Element4NumFilters.Value = 3
                    Case 15
                        Parameters.SetParameter(DocToUpdate.Document, "NumElementsWide", "4 ul")
                        Element1NumFilters.Value = 4
                        Element2NumFilters.Value = 3
                        Element3NumFilters.Value = 4
                        Element4NumFilters.Value = 4
                End Select
            Else
                'change something in one of the other files that isn't derived from the master part file.


            End If
        End If

    End Sub

    Sub UpdateFrontUpperFlange(ByVal MasterpartName As String)
        'MessageBox.Show("Hello World")
        'UpdateKeyWidthParameters()
        'need to include different offset as driven by GEK373 hole sketch blocks
        'need to refactor this to allow modification of the same or similar groups of parameters in different but related part files.
        ' it would be something like UpdateParams(Document, ElementNum,NumHoles,HoleOffset,PatternNumSlots,PatternStart
        If Parameters.GetParameter(DocToUpdate.Document, "MasterElement1NumFilters").Value = 3 Then
            'this line might need updating, I simply don't know as this stage.
            UpdateKeyParameters(MasterpartName, 1, 10, "188.75 mm", 5, "501.00 mm")
        Else
            'this line might need updating, I simply don't know as this stage.
            UpdateKeyParameters(MasterpartName, 1, 14, "193.75 mm", 8, "431.00 mm")
        End If
        If Parameters.GetParameter(DocToUpdate.Document, "MasterElement2NumFilters").Value = 3 Then
            UpdateKeyParameters(MasterpartName, 2, 10, "145.00 mm", 5, getPatternStart(2, MasterpartName))
        Else
            UpdateKeyParameters(MasterpartName, 2, 14, "150.00 mm", 8, getPatternStart(2, MasterpartName))
        End If
        If Parameters.GetParameter(DocToUpdate.Document, "MasterElement3NumFilters").Value = 3 Then
            UpdateKeyParameters(MasterpartName, 3, 10, "188.75 mm", 5, getPatternStart(3, MasterpartName))
        Else
            UpdateKeyParameters(MasterpartName, 3, 14, "150.00 mm", 8, getPatternStart(3, MasterpartName))
        End If
        If Parameters.GetParameter(DocToUpdate.Document, "MasterElement4NumFilters").Value = 3 Then
            UpdateKeyParameters(MasterpartName, 4, 10, "188.75 mm", 5, getPatternStart(4, MasterpartName))
        Else
            UpdateKeyParameters(MasterpartName, 4, 14, "193.75 mm", 8, getPatternStart(4, MasterpartName))
        End If
    End Sub



    Sub UpdateFrontBottomFlange(ByVal MasterpartName As String)
        'UpdateKeyWidthParameters()
        'need to add any width-specific parameter changes here.
        'check if this module is in the bottom-most position within a stack.
        If Parameters.GetParameter(DocToUpdate.Document, "FilterModuleInBottomPosition").Value = 1 Then
            '0 = shown
            Parameters.SetParameter(DocToUpdate.Document, "HideElement1BottomHole", 0)
            Parameters.SetParameter(DocToUpdate.Document, "HideElement2BottomHole", 0)
            If Parameters.GetParameter(DocToUpdate.Document, "NumElementsWide").Value = 4 Then
                Parameters.SetParameter(DocToUpdate.Document, "HideElement3BottomHole", 0)
            Else
                'need to hide this if we have less than 4 elements because the Clarcor guys cheated and deleted the face of the hole.
                Parameters.SetParameter(DocToUpdate.Document, "HideElement3BottomHole", 1)
            End If
        Else
            '1 = not shown
            Parameters.SetParameter(DocToUpdate.Document, "HideElement1BottomHole", 1)
            Parameters.SetParameter(DocToUpdate.Document, "HideElement2BottomHole", 1)
            Parameters.SetParameter(DocToUpdate.Document, "HideElement3BottomHole", 1)
        End If
        'Element 1 pattern 1 is independent of NumFilters
        UpdateKeyPatternParameters(MasterpartName, 1, 1, getPatternNumSlots(1, MasterpartName), getPatternStart(1, MasterpartName), "250.00 mm")
        If Parameters.GetParameter(DocToUpdate.Document, "MasterElement1NumFilters").Value = 3 Then
            UpdateKeyPatternParameters(MasterpartName, 1, 2, 2, "251.00 mm", "1500.00 mm")
        Else
            UpdateKeyPatternParameters(MasterpartName, 1, 2, 2, "251.00 mm", "2110.00 mm")
        End If
        'Element 2
        UpdateKeyPatternParameters(MasterpartName, 2, 1, getPatternNumSlots(2, MasterpartName), getPatternStart(2, MasterpartName), "250.00 mm")

        UpdateKeyPatternParameters(MasterpartName, 2, 2, 2, "200.00 mm", getPatternSpacing(2, MasterpartName))
        'Element 3
        UpdateKeyPatternParameters(MasterpartName, 3, 1, getPatternNumSlots(3, MasterpartName), getPatternStart(3, MasterpartName), "250.00 mm")

        UpdateKeyPatternParameters(MasterpartName, 3, 2, 2, "200.00 mm", getPatternSpacing(3, MasterpartName))
        'Element 4
        UpdateKeyPatternParameters(MasterpartName, 3, 1, getPatternNumSlots(4, MasterpartName), getPatternStart(4, MasterpartName), "250.00 mm")

        UpdateKeyPatternParameters(MasterpartName, 4, 2, 2, "200.00 mm", getPatternSpacing(4, MasterpartName))
        'debug from here as I simply copied this from the method above!
        If Parameters.GetParameter(DocToUpdate.Document, "MasterElement1NumFilters").Value = 3 Then
            'this line might need updating, I simply don't know as this stage.
            UpdateKeyParameters(MasterpartName, 1, 10, "188.75 mm", 5, "501.00 mm")
        Else
            'this line might need updating, I simply don't know as this stage.
            UpdateKeyParameters(MasterpartName, 1, 14, "193.75 mm", 8, "431.00 mm")
        End If
        If Parameters.GetParameter(DocToUpdate.Document, "MasterElement2NumFilters").Value = 3 Then
            UpdateKeyParameters(MasterpartName, 2, 10, "145.00 mm", 5, getPatternStart(2, MasterpartName))
        Else
            UpdateKeyParameters(MasterpartName, 2, 14, "150.00 mm", 8, getPatternStart(2, MasterpartName))
        End If
        If Parameters.GetParameter(DocToUpdate.Document, "MasterElement3NumFilters").Value = 3 Then
            UpdateKeyParameters(MasterpartName, 3, 10, "188.75 mm", 5, getPatternStart(3, MasterpartName))
        Else
            UpdateKeyParameters(MasterpartName, 3, 14, "150.00 mm", 8, getPatternStart(3, MasterpartName))
        End If
        If Parameters.GetParameter(DocToUpdate.Document, "MasterElement4NumFilters").Value = 3 Then
            UpdateKeyParameters(MasterpartName, 4, 10, "188.75 mm", 5, getPatternStart(4, MasterpartName))
        Else
            UpdateKeyParameters(MasterpartName, 4, 14, "193.75 mm", 8, getPatternStart(4, MasterpartName))
        End If

    End Sub



    Sub UpdateRearBottomFlange(ByVal MasterpartName As String)
        Select Case NumFiltersWide.Value
            Case 7

            Case 8
            Case 9
            Case 10
            Case 11
            Case 12
            Case 13
            Case 14
            Case 15

            Case Else

        End Select
    End Sub



    Sub UpdateInterBottomFlange(ByVal MasterpartName As String)
        Select Case NumFiltersWide.Value
            Case 7

            Case 8
            Case 9
            Case 10
            Case 11
            Case 12
            Case 13
            Case 14
            Case 15

            Case Else

        End Select
    End Sub

    Sub UpdateRearUpperFlange(ByVal MasterpartName As String)
        Select Case NumFiltersWide.Value
            Case 7

            Case 8
            Case 9
            Case 10
            Case 11
            Case 12
            Case 13
            Case 14
            Case 15

            Case Else

        End Select
    End Sub


    Sub UpdateInterUpperFlange(ByVal MasterpartName As String)
        Select Case NumFiltersWide.Value
            Case 7

            Case 8
            Case 9
            Case 10
            Case 11
            Case 12
            Case 13
            Case 14
            Case 15

            Case Else

        End Select
    End Sub
#End Region

#Region "Vertical Members"
    Private Sub UpdateStage1LeftVertical(ByVal MasterpartName As String)
        If "parameter is bottom member" Then

        ElseIf "the other option" Then

        End If
    End Sub
#End Region


    ''' <summary>
    ''' Grabs our pattern start value based on masterpartname and elementnum
    ''' </summary>
    ''' <param name="ElementNum"></param>
    ''' <param name="masterpartname"></param>
    ''' <returns></returns>
    Private Function getPatternStart(ByVal ElementNum As Integer, ByVal masterpartname As String) As String
        Select Case masterpartname
            Case "Master-FrontUpperFlange"
                ' I am aware that these could be combined. 2016-10-03 AF
                Select Case ElementNum
                    Case 1
                        Select Case NumFiltersWide.Value
                            Case 7
                                Return "250.00 mm"
                            Case 8
                                Return "180.00 mm"
                            Case 9, 11, 13, 15
                                Return "215.00 mm"
                            Case 10
                                Return "145.00 mm"
                            Case 12, 14
                                Return "270.00 mm"
                            Case Else
                                Return "133.7 mm" 'because l33t 8-)
                        End Select
                    Case 2
                        Select Case NumFiltersWide.Value
                            Case 7
                                Return "250.00 mm"
                            Case 8
                                Return "180.00 mm"
                            Case 9, 11, 13, 15
                                Return "215.00 mm"
                            Case 10
                                Return "145.00 mm"
                            Case 12, 14
                                Return "270.00 mm"
                            Case Else
                                Return "133.7 mm" 'because l33t 8-)
                        End Select
                    Case 3
                        Select Case NumFiltersWide.Value
                            Case 9, 10
                                Return "250.00 mm"
                            Case 11, 12
                                Return "180.00 mm"
                            Case 13, 15
                                Return "145.00 mm"
                            Case 14
                                Return "270.00 mm"
                            Case Else
                                Return "133.7 mm" 'because l33t 8-)
                        End Select
                    Case 4
                        Select Case NumFiltersWide.Value
                            Case 13, 14
                                Return "450.00 mm"
                            Case 15
                                Return "380.00 mm"
                            Case Else
                                Return "133.7 mm" 'because l33t 8-)
                        End Select
                End Select
            Case "Master-FrontBottomFlange"
                Select Case ElementNum
                    Case 1
                        Select Case NumFiltersWide.Value
                            Case 7, 8, 11, 12, 15
                                Return "431.00 mm"
                            Case 9
                                Return "501.00 mm"
                            Case 10, 13, 14
                                Return "251.00 mm"
                            Case Else
                                Return "133.7 mm" 'because l33t 8-)
                        End Select
                    Case 2
                        Select Case NumFiltersWide.Value
                            Case 7
                                Return "450.00 mm"
                            Case 8
                                Return "380.00 mm"
                            Case 9, 11, 13, 15
                                Return "415.00 mm"
                            Case 10, 12
                                Return "345.00 mm"
                            Case 14
                                Return "470.00 mm"
                            Case Else
                                Return "133.7 mm" 'because l33t 8-)
                        End Select
                    Case 3
                        Select Case NumFiltersWide.Value
                            Case 9, 10
                                Return "450.00 mm"
                            Case 11, 12
                                Return "380.00 mm"
                            Case 13, 14, 15
                                Return "345.00 mm"
                            Case Else
                                Return "133.7 mm" 'because l33t 8-)
                        End Select
                    Case 4
                        Select Case NumFiltersWide.Value
                            Case 13, 14
                                Return "450.00 mm"
                            Case 15
                                Return "380.00 mm"
                            Case Else
                                Return "133.7 mm" 'because l33t 8-)
                        End Select
                End Select
            Case Else
                Throw New NotImplementedException
        End Select
    End Function


    Private Function getPatternSpacing(ByVal ElementNum As Integer, ByVal masterpartname As String) As String
        Select Case masterpartname
            Case "Master-FrontUpperFlange"

            Case """Master-FrontBottomFlange"
                Select Case ElementNum
                    Case 1

                    Case 2
                        Select Case NumFiltersWide.Value
                            Case 7
                                Return "1500.00 mm"
                            Case 8
                                Return "2110.00 mm"
                            Case 9, 11, 13
                                Return "1430.00 mm"
                            Case 10, 12, 14
                                Return "2040.00 mm"
                            Case 15
                                Return "1430.00 mm"
                        End Select
                    Case 3
                        Select Case NumFiltersWide.Value
                            Case 7, 8
                                Return "1337.00 mm * 1.5 ul" ' because we can't have this being 0
                            Case 9, 10
                                Return "1500.00 mm"
                            Case 11, 12
                                Return "2100.00 mm"
                            Case 13, 14, 15
                                Return "2040.00 mm"
                        End Select
                    Case 4
                        Select Case NumFiltersWide.Value
                            Case 13, 14
                                Return "1500.00 mm"
                            Case 15
                                Return "2110.00 mm"
                            Case Else
                                Return "1337.00 mm * 1.5 ul" ' because we can't have this being 0
                        End Select
                End Select
            Case Else

        End Select
    End Function

    ''' <summary>
    ''' A global variation of the individual getPattern{#}Numslots method.
    ''' </summary>
    ''' <param name="ElementNum"></param>
    ''' <param name="masterpartname"></param>
    ''' <returns></returns>
    Private Function getPatternNumSlots(ByVal ElementNum As Integer, ByVal masterpartname As String) As Integer
        Select Case masterpartname
            Case "Master-FrontUpperFlange"
                Select Case ElementNum
                    Case 1

                    Case 2

                    Case 3

                    Case 4

                End Select
            Case "Master-FrontBottomFlange"
                Select Case ElementNum
                    Case 1
                        Select Case NumFiltersWide.Value
                            Case 7, 8, 11, 12, 15
                                Return 8
                            Case 9
                                Return 5
                            Case 10, 14
                                Return 6
                            Case 13
                                Return 7
                            Case Else
                                Return 0
                        End Select
                    Case 2
                        Select Case NumFiltersWide.Value
                            Case 7, 9, 11, 13, 15
                                Return 5
                            Case 8, 10, 12
                                Return 8
                            Case 14
                                Return 7
                        End Select
                    Case 3
                        Select Case NumFiltersWide.Value
                            Case 7, 8, 9, 10
                                Return 5
                            Case Else
                                Return 8
                        End Select
                    Case 4
                        Select Case NumFiltersWide.Value
                            Case 13, 14
                                Return 5
                            Case 15
                                Return 8
                            Case Else
                                Return 7 'just returning a default value even though the pattern that depends upon it will be suppressed.
                        End Select
                End Select
            Case Else
        End Select
    End Function




#End Region
End Class

Public Class ModuleElement
    Public Property ElementPosition As Integer
        Get
            Return _elementPosition
        End Get
        Set(value As Integer)
            _elementPosition = value
        End Set
    End Property

    Private _elementPosition As Integer

    Public Sub New(ElementPosn As Integer)
        ElementPosition = ElementPosn
    End Sub

    Private m_Pattern1Start As String
    Public Property Pattern1Start() As String
        Get
            Return m_Pattern1Start
        End Get
        Set(ByVal value As String)
            m_Pattern1Start = value
        End Set
    End Property

    Private m_Pattern1Spacing As String
    Public Property Pattern1Spacing() As String
        Get
            Return m_Pattern1Spacing
        End Get
        Set(ByVal value As String)
            m_Pattern1Spacing = value
        End Set
    End Property

    Private m_Pattern2Start As String
    Public Property Pattern2Start() As String
        Get
            Return m_Pattern2Start
        End Get
        Set(ByVal value As String)
            m_Pattern2Start = value
        End Set
    End Property

    Private m_Pattern2Spacing As String
    Public Property Pattern2Spacing() As String
        Get
            Return m_Pattern2Spacing
        End Get
        Set(ByVal value As String)
            m_Pattern2Spacing = value
        End Set
    End Property

    Private m_NumFilters As String
    Public Property NumFilters() As String
        Get
            Return m_NumFilters
        End Get
        Set(ByVal value As String)
            m_NumFilters = value
        End Set
    End Property

    Private m_ElementMasterName As String
    Public Property ElementMasterName() As String
        Get
            Return m_ElementMasterName
        End Get
        Set(ByVal value As String)
            m_ElementMasterName = "MasterElement" & ElementPosition.ToString() & "NumFilters"
        End Set
    End Property
    'Public MustOverride Function GetDefaultPatternSpacing() As List(Of String)
    'Public MustOverride Function GetDefaultPatternStartPosition() As List(Of String)
    'Public MustOverride Function GetDefaultPatternNumSlots() As Integer
    Public Function DefaultPattern1Spacing() As Integer

        Return 0
    End Function

    Public Function DefaultSlotPatternSpacing() As String
        Select Case ElementPosition
            Case 1

            Case 2

            Case 3

            Case 4


            Case Else

        End Select

        Return 0
    End Function
End Class

Public MustInherit Class ModuleStage
    Private m_StageDepth As String
    Public Property Depth() As String
        Get
            Return m_StageDepth
        End Get
        Set(ByVal value As String)
            m_StageDepth = value
        End Set
    End Property
    Private m_NumBottomSlots As String
    Public Property NumBottomSlots() As String
        Get
            Return m_NumBottomSlots
        End Get
        Set(ByVal value As String)
            m_NumBottomSlots = value
        End Set
    End Property

    Private m_NumSideSlots As String
    Public Property NumSideSlots() As String
        Get
            Return m_NumSideSlots
        End Get
        Set(ByVal value As String)
            m_NumSideSlots = value
        End Set
    End Property
End Class

Public Class TwoStageFiltration
    Inherits ModuleStage

End Class

Public Class ThreeStageFiltration
    Inherits ModuleStage

End Class

Public Class FourStageFiltration
    Inherits ModuleStage

End Class
Public Class ThreeWideElement
    Inherits ModuleElement

    Public Sub New(ElementPosn As Integer)
        MyBase.New(ElementPosn)
    End Sub
End Class

Public Class FourWideElement
    Inherits ModuleElement

    Public Sub New(ElementPosn As Integer)
        MyBase.New(ElementPosn)
    End Sub
End Class

Public Class FilterModule

    Public Property MaxNumFiltersWide As Integer = 15

    Private m_NumFiltersThisModule As Integer
    Public Property NumFiltersThisModule() As Integer
        Get
            Return m_NumFiltersThisModule
        End Get
        Set(ByVal value As Integer)
            m_NumFiltersThisModule = value
        End Set
    End Property
    Private m_IsLeftModule As Boolean
    Public Property IsLeftModule() As Boolean
        Get
            Return m_IsLeftModule
        End Get
        Set(ByVal value As Boolean)
            m_IsLeftModule = value
        End Set
    End Property
    Private m_ElementOne As ModuleElement
    Public Property ElementOne() As ModuleElement
        Get
            Return m_ElementOne
        End Get
        Set(ByVal value As ModuleElement)
            m_ElementOne = value
        End Set
    End Property
    Private m_Modules As List(Of ModuleElement)
    Public Property Modules() As List(Of ModuleElement)
        Get
            Return m_Modules
        End Get
        Set(ByVal value As List(Of ModuleElement))
            m_Modules = value
        End Set
    End Property
    Private m_ElementTwo As ModuleElement
    Public Property ElementTwo() As ModuleElement
        Get
            Return m_ElementTwo
        End Get
        Set(ByVal value As ModuleElement)
            m_ElementTwo = value
        End Set
    End Property

    Private m_ElementThree As ModuleElement
    Public Property ElementThree() As ModuleElement
        Get
            Return m_ElementThree
        End Get
        Set(ByVal value As ModuleElement)
            m_ElementThree = value
        End Set
    End Property

    Private m_ElementFour As ModuleElement
    Public Property ElementFour() As ModuleElement
        Get
            Return m_ElementFour
        End Get
        Set(ByVal value As ModuleElement)
            m_ElementFour = value
        End Set
    End Property

    Private m_NumElementsWide As Integer
    Public Property NumElementsWide() As Integer
        Get
            Return m_NumElementsWide
        End Get
        Set(ByVal value As Integer)
            m_NumElementsWide = value
        End Set
    End Property

    Private m_NumFiltrationStages As Integer
    Public Property NumFiltrationStages() As Integer
        Get
            Return m_NumFiltrationStages
        End Get
        Set(ByVal value As Integer)
            m_NumFiltrationStages = value
        End Set
    End Property

    Private m_StageOne As ModuleStage
    Public Property StageOne() As ModuleStage
        Get
            Return m_StageOne
        End Get
        Set(ByVal value As ModuleStage)
            m_StageOne = value
        End Set
    End Property


    Private m_StageTwo As ModuleStage
    Public Property StageTwo() As ModuleStage
        Get
            Return m_StageTwo
        End Get
        Set(ByVal value As ModuleStage)
            m_StageTwo = value
        End Set
    End Property

    Private m_NumFiltersHigh As Integer
    Public Property NumFiltersHigh() As Integer
        Get
            Return m_NumFiltersHigh
        End Get
        Set(ByVal value As Integer)
            m_NumFiltersHigh = value
        End Set
    End Property

    Private m_NumFiltersWide As Integer
    Public Property NumFiltersWide() As Integer
        Get
            Return m_NumFiltersWide
        End Get
        Set(ByVal value As Integer)
            m_NumFiltersWide = value
        End Set
    End Property



    Public Sub New(v1 As String, v2 As String, v3 As String, v4 As String)
        NumFiltersWide = Convert.ToInt32(v1.Replace(" ul", String.Empty))
        NumFiltersHigh = Convert.ToInt32(v2.Replace(" ul", String.Empty))
        NumFiltrationStages = Convert.ToInt32(v3.Replace(" ul", String.Empty))
        Modules = New List(Of ModuleElement)
        'should be set programatically?
        'NumElementsWide = Convert.ToInt32(v4.Replace(" ul", String.Empty))
        Select Case NumFiltersWide
            Case 7
                ElementOne = New FourWideElement(ElementPosn:=1)
                ElementTwo = New ThreeWideElement(ElementPosn:=2)
                Modules.Add(ElementOne)
                Modules.Add(ElementTwo)
            Case 8
                ElementOne = New FourWideElement(ElementPosn:=1)
                ElementTwo = New FourWideElement(ElementPosn:=2)
                Modules.Add(ElementOne)
                Modules.Add(ElementTwo)
            Case 9
                ElementOne = New ThreeWideElement(ElementPosn:=1)
                ElementTwo = New ThreeWideElement(ElementPosn:=2)
                ElementThree = New ThreeWideElement(ElementPosn:=3)
                Modules.Add(ElementOne)
                Modules.Add(ElementTwo)
                Modules.Add(ElementThree)
            Case 10
                ElementOne = New ThreeWideElement(ElementPosn:=1)
                ElementTwo = New FourWideElement(ElementPosn:=2)
                ElementThree = New ThreeWideElement(ElementPosn:=3)
                Modules.Add(ElementOne)
                Modules.Add(ElementTwo)
                Modules.Add(ElementThree)
            Case 11
                ElementOne = New FourWideElement(ElementPosn:=1)
                ElementTwo = New ThreeWideElement(ElementPosn:=2)
                ElementThree = New FourWideElement(ElementPosn:=3)
                Modules.Add(ElementOne)
                Modules.Add(ElementTwo)
                Modules.Add(ElementThree)
            Case 12
                ElementOne = New FourWideElement(ElementPosn:=1)
                ElementTwo = New FourWideElement(ElementPosn:=2)
                ElementThree = New FourWideElement(ElementPosn:=3)
                Modules.Add(ElementOne)
                Modules.Add(ElementTwo)
                Modules.Add(ElementThree)
            Case 13
                ElementOne = New ThreeWideElement(ElementPosn:=1)
                ElementTwo = New ThreeWideElement(ElementPosn:=2)
                ElementThree = New FourWideElement(ElementPosn:=3)
                ElementFour = New ThreeWideElement(ElementPosn:=4)
                Modules.Add(ElementOne)
                Modules.Add(ElementTwo)
                Modules.Add(ElementThree)
                Modules.Add(ElementFour)
            Case 14
                ElementOne = New ThreeWideElement(ElementPosn:=1)
                ElementTwo = New FourWideElement(ElementPosn:=2)
                ElementThree = New FourWideElement(ElementPosn:=3)
                ElementFour = New ThreeWideElement(ElementPosn:=4)
                Modules.Add(ElementOne)
                Modules.Add(ElementTwo)
                Modules.Add(ElementThree)
                Modules.Add(ElementFour)
            Case 15
                ElementOne = New FourWideElement(ElementPosn:=1)
                ElementTwo = New ThreeWideElement(ElementPosn:=2)
                ElementThree = New FourWideElement(ElementPosn:=3)
                ElementFour = New FourWideElement(ElementPosn:=4)
                Modules.Add(ElementOne)
                Modules.Add(ElementTwo)
                Modules.Add(ElementThree)
                Modules.Add(ElementFour)
        End Select
        Select Case NumFiltrationStages
            Case 2
                StageOne = New TwoStageFiltration()
            Case 3

            Case 4

        End Select

    End Sub


End Class
#Region "Helper classes"



Public Class iProperties

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
                prop = customPropSet.Add(PropertyValue, PropertyName)
            Else
                ' Change the value of the existing property.
                prop.Value = PropertyValue
            End If
        Else
            customproperty = prop.Value
        End If
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
