Imports Inventor

Public MustInherit Class SheetMetal
    Private f_currentThickness As Double
    Private f_currentSheetMetalStyle As SheetMetalStyle

    Public Property CurrentThickness As Double
        Get
            Return f_currentThickness
        End Get
        Set(value As Double)
            f_currentThickness = value
        End Set
    End Property

    Public Property CurrentSheetMetalStyle As SheetMetalStyle
        Get
            Return GetCurrentSheetMetalStyle(SheetMetalDoc)
        End Get
        Set(value As SheetMetalStyle)
            f_currentSheetMetalStyle = value
        End Set
    End Property

    Private f_sheetMetalDoc As PartDocument
    Public Property SheetMetalDoc() As PartDocument
        Get
            Return f_sheetMetalDoc
        End Get
        Set(ByVal value As PartDocument)
            f_sheetMetalDoc = value
        End Set
    End Property


    ''' <summary>
    ''' Gets the currently active sheet metal style.
    ''' </summary>
    ''' <param name="Doc">The Inventor document to query</param>
    ''' <returns></returns>
    Public Function GetCurrentSheetMetalStyle(ByVal Doc As Document) As SheetMetalStyle
        'only works on part files that have a sheet metal definition
        SheetMetalDoc = Doc
        Dim sheetmetalCompdef As SheetMetalComponentDefinition = SheetMetalDoc.ComponentDefinition
        If Not sheetmetalCompdef Is Nothing Then
            Return sheetmetalCompdef.ActiveSheetMetalStyle
        Else
            Return Nothing
        End If
    End Function

    ''' <summary>
    ''' This should let us set the currently active sheet metal style.
    ''' </summary>
    ''' <example>https://forums.autodesk.com/t5/inventor-customization/is-there-a-way-to-change-the-sheet-metal-rule/td-p/2568127</example>
    ''' <param name="StyleName">The style name we wish to set as active</param>
    ''' <returns></returns>
    Public Function SetorCreateCurrentSheetMetalStyle(ByVal Doc As Document, ByVal StyleName As String) As SheetMetalStyle
        If Not StyleName = String.Empty Then
            SheetMetalDoc = Doc
            Dim sheetmetalCompdef As SheetMetalComponentDefinition = SheetMetalDoc.ComponentDefinition
            If Not sheetmetalCompdef Is Nothing Then
                Dim shtmetalstyle As SheetMetalStyle = (From style As SheetMetalStyle In sheetmetalCompdef.SheetMetalStyles
                                                        Where style.Name = StyleName
                                                        Select style).First()
                If Not shtmetalstyle Is Nothing Then
                    'the style to change to exists in the document.
                    sheetmetalCompdef.SheetMetalStyles.Item(StyleName).Activate()
                    Return shtmetalstyle
                Else
                    ' it doesn't exist and should be created.
                    Return Nothing
                End If

            Else
                Return Nothing
            End If
        Else
            Return Nothing
        End If
    End Function

    ''' <summary>
    ''' Gets or sets the active kFactor based on a string input
    ''' </summary>
    ''' <param name="kFactor"></param>
    ''' <returns></returns>
    Public Function GetorSetActiveKFactor(ByVal Doc As Document, ByVal kFactor As String) As String
        SheetMetalDoc = Doc
        Dim sheetmetalCompdef As SheetMetalComponentDefinition = SheetMetalDoc.ComponentDefinition
        If Not sheetmetalCompdef Is Nothing Then
            Dim shtmetalstyle As SheetMetalStyle = sheetmetalCompdef.ActiveSheetMetalStyle
            If Not kFactor.Equals(String.Empty) Then
                shtmetalstyle.UnfoldMethod.kFactor = kFactor
                Return shtmetalstyle.UnfoldMethod.kFactor
            Else 'return the current
                Return shtmetalstyle.UnfoldMethod.kFactor
            End If
        Else
            Return String.Empty
        End If
    End Function
End Class
