﻿Imports Inventor
Namespace Power_Pack_For_Inventor_AddIn
    ''' <summary>
    ''' All of these functions will work with the currently active document inside of Inventor.
    ''' </summary>
    Public Class Parameters
#Region "Parameters"
        ''' <summary>
        ''' UNTESTED 2016-05-24 AF
        ''' Sets a string parameter value
        ''' </summary>
        ''' <param name="ParameterName"></param>
        ''' <param name="ParameterValue"></param>
        Public Shared Sub SetParameter(ByVal ParameterName As String, ByVal ParameterValue As String)
            ' Get the Parameters object. Assumes a part or assembly document is active.
            Dim oParameters As Inventor.Parameters = g_inventorApplication.ActiveDocument.ComponentDefinition.Parameters

            ' Get the parameter named "Length".
            Dim oLengthParam As Inventor.Parameter = oParameters.Item(ParameterName)

            ' Change the equation of the parameter.
            oLengthParam.Expression = ParameterValue

            ' Update the document.
            If Power_Pack_For_Inventor_AddIn.UpdateAfterEachParameterChange Then
                g_inventorApplication.ActiveDocument.Update()
            End If
        End Sub
        ''' <summary>
        ''' UNTESTED 2016-05-24 AF
        ''' Sets a number parameter value
        ''' </summary>
        ''' <param name="ParameterName"></param>
        ''' <param name="ParameterValue"></param>
        Public Shared Sub SetParameter(ByVal ParameterName As String, ByVal ParameterValue As Double)
            ' Get the Parameters object. Assumes a part or assembly document is active.
            Dim oParameters As Inventor.Parameters = g_inventorApplication.ActiveDocument.ComponentDefinition.Parameters

            ' Get the parameter named "Length".
            Dim oLengthParam As Inventor.Parameter = oParameters.Item(ParameterName)

            ' Change the equation of the parameter.
            oLengthParam.Expression = ParameterValue

            ' Update the document.
            If Power_Pack_For_Inventor_AddIn.UpdateAfterEachParameterChange Then
                g_inventorApplication.ActiveDocument.Update()
            End If
        End Sub
        ''' <summary>
        ''' UNTESTED 2016-05-24 AF
        ''' Sets a true/false parameter value
        ''' </summary>
        ''' <param name="ParameterName"></param>
        ''' <param name="ParameterValue"></param>
        Public Shared Sub SetParameter(ByVal ParameterName As String, ByVal ParameterValue As Boolean)
            ' Get the Parameters object. Assumes a part or assembly document is active.
            Dim oParameters As Inventor.Parameters = g_inventorApplication.ActiveDocument.ComponentDefinition.Parameters

            ' Get the parameter named "Length".
            Dim oLengthParam As Inventor.Parameter = oParameters.Item(ParameterName)

            ' Change the equation of the parameter.
            oLengthParam.Expression = ParameterValue

            ' Update the document.
            If Power_Pack_For_Inventor_AddIn.UpdateAfterEachParameterChange Then
                g_inventorApplication.ActiveDocument.Update()
            End If
        End Sub
        ''' <summary>
        ''' UNTESTED 2016-05-24 AF
        ''' Sets a Date Parameter Value
        ''' </summary>
        ''' <param name="ParameterName"></param>
        ''' <param name="ParameterValue"></param>
        Public Shared Sub SetParameter(ByVal ParameterName As String, ByVal ParameterValue As DateTime)
            ' Get the Parameters object. Assumes a part or assembly document is active.
            Dim oParameters As Inventor.Parameters = g_inventorApplication.ActiveDocument.ComponentDefinition.Parameters

            ' Get the parameter named "Length".
            Dim oLengthParam As Inventor.Parameter = oParameters.Item(ParameterName)

            ' Change the equation of the parameter.
            oLengthParam.Expression = ParameterValue

            ' Update the document.
            If Power_Pack_For_Inventor_AddIn.UpdateAfterEachParameterChange Then
                g_inventorApplication.ActiveDocument.Update()
            End If
        End Sub

        Public Shared Function GetParameter(ByVal ParamName As String) As Inventor.Parameter
            Dim oParameters As Inventor.Parameters = g_inventorApplication.ActiveDocument.ComponentDefinition.Parameters
            If oParameters(ParamName).ParameterType = ParameterTypeEnum.kUserParameter Then
                Return GetUserParameter(ParamName)
            ElseIf oParameters(ParamName).ParameterType = ParameterTypeEnum.kReferenceParameter Then
                Return GetReferenceParameter(ParamName)
            ElseIf oParameters(ParamName).ParameterType = ParameterTypeEnum.kModelParameter Then
                Return GetModelParameter(ParamName)
            ElseIf oParameters(ParamName).ParameterType = ParameterTypeEnum.kDerivedParameter Then
                Return GetDerivedParameter(ParamName)
            ElseIf oParameters(ParamName).ParameterType = ParameterTypeEnum.kTableParameter Then
                Throw New NotSupportedException()
            End If
            Return Nothing
        End Function

        ''' <summary>
        ''' Gets the object of a parameter by name
        ''' </summary>
        ''' <param name="ParameterName"></param>
        ''' <returns></returns>
        Public Shared Function GetUserParameter(ByVal ParameterName As String) As Inventor.UserParameter
            Dim oParameters As Inventor.Parameters = g_inventorApplication.ActiveDocument.ComponentDefinition.Parameters
            Return oParameters.Item(ParameterName)
        End Function

        ''' <summary>
        ''' Gets the object of a reference parameter by name
        ''' </summary>
        ''' <param name="ParameterName"></param>
        ''' <returns></returns>
        Public Shared Function GetReferenceParameter(ByVal ParameterName As String) As Inventor.ReferenceParameter
            Dim oParameters As Inventor.Parameters = g_inventorApplication.ActiveDocument.ComponentDefinition.Parameters
            Return oParameters.Item(ParameterName)
        End Function

        ''' <summary>
        ''' Gets the object of a model parameter by name
        ''' </summary>
        ''' <param name="ParameterName"></param>
        ''' <returns></returns>
        Public Shared Function GetModelParameter(ByVal ParameterName As String) As Inventor.ModelParameter
            Dim oParameters As Inventor.Parameters = g_inventorApplication.ActiveDocument.ComponentDefinition.Parameters
            Return oParameters.Item(ParameterName)
        End Function

        ''' <summary>
        ''' Gets the object of a derived parameter by name
        ''' </summary>
        ''' <param name="ParameterName"></param>
        ''' <returns></returns>
        Public Shared Function GetDerivedParameter(ByVal ParameterName As String) As Inventor.DerivedParameter
            Dim oParameters As Inventor.Parameters = g_inventorApplication.ActiveDocument.ComponentDefinition.Parameters
            Return oParameters.Item(ParameterName)
        End Function


#End Region

    End Class
End Namespace