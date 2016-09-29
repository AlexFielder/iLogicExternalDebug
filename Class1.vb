Imports System.Windows.Forms
Imports Inventor
Imports Autodesk.iLogic.Interfaces

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

        Select Case iProperties.SetorCreateCustomiProperty(DocToUpdate.Document, "BasePartNumber")
            Case "Master-FrontUpperFlange"
                UpdateFrontUpperFlange()
            Case "Master-FrontBottomFlange"
                UpdateFrontBottomFlange()
            Case "Master-RearBottomFlange"
                UpdateRearBottomFlange()
            Case "Master-InterBottomFlange"
                UpdateInterBottomFlange()
            Case "Master-RearUpperFlange"
                UpdateRearUpperFlange()
            Case "Master-InterUpperFlange"
                UpdateInterUpperFlange()

            Case "FilterModuleMaster" ' our master part file so this better not break anything!
                UpdateFilterModuleMaster()
            Case "Add more as required"

            Case Else

        End Select


    End Sub
    'Key Parameters
    Public NumFiltersWide As Inventor.Parameter = Nothing
    Public NumFiltersHigh As Inventor.Parameter = Nothing
    Public NumFiltrationStages As Inventor.Parameter = Nothing

    Sub UpdateFilterModuleMaster()
        'MessageBox.Show("Hello World")
        UpdateKeyParameters()

    End Sub

    Sub UpdateFrontUpperFlange()
        'MessageBox.Show("Hello World")
        UpdateKeyParameters()

    End Sub

    Private Sub UpdateKeyParameters()
        If TypeOf DocToUpdate.Document Is PartDocument Then
            Dim derivedpartcheck As PartDocument = DocToUpdate.Document
            If derivedpartcheck.ComponentDefinition.ReferenceComponents.DerivedPartComponents.Count = 0 Then
                Select Case NumFiltersWide.Value
                    Case 7
                        Parameters.SetParameter(DocToUpdate.Document, "NumElementsWide", "2 ul")
                    Case 8
                        Parameters.SetParameter(DocToUpdate.Document, "NumElementsWide", "2 ul")
                    Case 9
                        Parameters.SetParameter(DocToUpdate.Document, "NumElementsWide", "3 ul")
                    Case 10
                        Parameters.SetParameter(DocToUpdate.Document, "NumElementsWide", "3 ul")
                    Case 11
                        Parameters.SetParameter(DocToUpdate.Document, "NumElementsWide", "3 ul")
                    Case 12
                        Parameters.SetParameter(DocToUpdate.Document, "NumElementsWide", "3 ul")
                    Case 13
                        Parameters.SetParameter(DocToUpdate.Document, "NumElementsWide", "4 ul")
                    Case 14
                        Parameters.SetParameter(DocToUpdate.Document, "NumElementsWide", "4 ul")
                    Case 15
                        Parameters.SetParameter(DocToUpdate.Document, "NumElementsWide", "4 ul")
                End Select
            Else
                Dim Element1NumFilters As Parameter = Parameters.GetParameter(DocToUpdate.Document, "Element1NumFilters")
                Dim Element2NumFilters As Parameter = Parameters.GetParameter(DocToUpdate.Document, "Element2NumFilters")
                Dim Element3NumFilters As Parameter = Parameters.GetParameter(DocToUpdate.Document, "Element3NumFilters")
                Dim Element4NumFilters As Parameter = Parameters.GetParameter(DocToUpdate.Document, "Element4NumFilters")
                Select Case NumFiltersWide.Value
                    Case 7
                        Element1NumFilters.Value = 4
                        Element2NumFilters.Value = 3
                    Case 8
                        Element1NumFilters.Value = 4
                        Element2NumFilters.Value = 4
                    Case 9
                        Element1NumFilters.Value = 3
                        Element2NumFilters.Value = 3
                        Element3NumFilters.Value = 3
                    Case 10
                        Element1NumFilters.Value = 3
                        Element2NumFilters.Value = 4
                        Element3NumFilters.Value = 3
                    Case 11
                        Element1NumFilters.Value = 4
                        Element2NumFilters.Value = 3
                        Element3NumFilters.Value = 4
                    Case 12
                        Element1NumFilters.Value = 4
                        Element2NumFilters.Value = 4
                        Element3NumFilters.Value = 4
                    Case 13
                        Element1NumFilters.Value = 3
                        Element2NumFilters.Value = 3
                        Element3NumFilters.Value = 4
                        Element4NumFilters.Value = 3
                    Case 14
                        Element1NumFilters.Value = 3
                        Element2NumFilters.Value = 4
                        Element3NumFilters.Value = 4
                        Element4NumFilters.Value = 3
                    Case 15
                        Element1NumFilters.Value = 4
                        Element2NumFilters.Value = 3
                        Element3NumFilters.Value = 4
                        Element4NumFilters.Value = 4

                    Case Else

                End Select
            End If
        End If

    End Sub

    Sub UpdateFrontBottomFlange()
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


    Sub UpdateRearBottomFlange()
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



    Sub UpdateInterBottomFlange()
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

    Sub UpdateRearUpperFlange()
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


    Sub UpdateInterUpperFlange()
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
End Class

Public Class FilterModule
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
    Private m_Element1FilterCount As Integer
    Public Property Element1FilterCount() As Integer
        Get
            Return m_Element1FilterCount
        End Get
        Set(ByVal value As Integer)
            m_Element1FilterCount = value
        End Set
    End Property

    Private m_Element2FilterCount As String
    Public Property Element2FilterCount() As String
        Get
            Return m_Element2FilterCount
        End Get
        Set(ByVal value As String)
            m_Element2FilterCount = value
        End Set
    End Property
    Private m_Element3FilterCount As String
    Public Property Element3FilterCount() As String
        Get
            Return m_Element3FilterCount
        End Get
        Set(ByVal value As String)
            m_Element3FilterCount = value
        End Set
    End Property
    Private m_Element4FilterCount As String
    Public Property Element4FilterCount() As String
        Get
            Return m_Element4FilterCount
        End Get
        Set(ByVal value As String)
            m_Element4FilterCount = value
        End Set
    End Property
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
