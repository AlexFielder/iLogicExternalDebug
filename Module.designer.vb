'' ------------------------------------------------------------------------------
''  <auto-generated>
''    Generated by Xsd2Code. Version 3.4.0.32989
''    <NameSpace>Module.vb</NameSpace><Collection>List</Collection><codeType>VisualBasic</codeType><EnableDataBinding>False</EnableDataBinding><EnableLazyLoading>False</EnableLazyLoading><TrackingChangesEnable>False</TrackingChangesEnable><GenTrackingClasses>False</GenTrackingClasses><HidePrivateFieldInIDE>False</HidePrivateFieldInIDE><EnableSummaryComment>True</EnableSummaryComment><VirtualProp>False</VirtualProp><IncludeSerializeMethod>True</IncludeSerializeMethod><UseBaseClass>True</UseBaseClass><GenBaseClass>True</GenBaseClass><GenerateCloneMethod>False</GenerateCloneMethod><GenerateDataContracts>False</GenerateDataContracts><CodeBaseTag>Net20</CodeBaseTag><SerializeMethodName>Serialize</SerializeMethodName><DeserializeMethodName>Deserialize</DeserializeMethodName><SaveToFileMethodName>SaveToFile</SaveToFileMethodName><LoadFromFileMethodName>LoadFromFile</LoadFromFileMethodName><GenerateXMLAttributes>True</GenerateXMLAttributes><OrderXMLAttrib>False</OrderXMLAttrib><EnableEncoding>False</EnableEncoding><AutomaticProperties>False</AutomaticProperties><GenerateShouldSerialize>False</GenerateShouldSerialize><DisableDebug>False</DisableDebug><PropNameSpecified>Default</PropNameSpecified><Encoder>UTF8</Encoder><CustomUsings></CustomUsings><ExcludeIncludedTypes>False</ExcludeIncludedTypes><EnableInitializeFields>True</EnableInitializeFields>
''  </auto-generated>
'' ------------------------------------------------------------------------------
Imports System
Imports System.Diagnostics
Imports System.Xml.Serialization
Imports System.Collections
Imports System.Xml.Schema
Imports System.ComponentModel
Imports System.IO
Imports System.Text
Imports System.Collections.Generic

Namespace [Module].vb
    
    #Region "Base entity class"
    Partial Public Class EntityBase(Of T)
        
        Private Shared sSerializer As System.Xml.Serialization.XmlSerializer
        
        Private Shared ReadOnly Property Serializer() As System.Xml.Serialization.XmlSerializer
            Get
                If (sSerializer Is Nothing) Then
                    sSerializer = New System.Xml.Serialization.XmlSerializer(GetType(T))
                End If
                Return sSerializer
            End Get
        End Property
        
        #Region "Serialize/Deserialize"
        '''<summary>
        '''Serializes current EntityBase object into an XML document
        '''</summary>
        '''<returns>string XML value</returns>
        Public Overridable Function Serialize() As String
            Dim streamReader As System.IO.StreamReader = Nothing
            Dim memoryStream As System.IO.MemoryStream = Nothing
            Try 
                memoryStream = New System.IO.MemoryStream
                Serializer.Serialize(memoryStream, Me)
                memoryStream.Seek(0, System.IO.SeekOrigin.Begin)
                streamReader = New System.IO.StreamReader(memoryStream)
                Return streamReader.ReadToEnd
            Finally
                If (Not (streamReader) Is Nothing) Then
                    streamReader.Dispose
                End If
                If (Not (memoryStream) Is Nothing) Then
                    memoryStream.Dispose
                End If
            End Try
        End Function
        
        '''<summary>
        '''Deserializes workflow markup into an EntityBase object
        '''</summary>
        '''<param name="xml">string workflow markup to deserialize</param>
        '''<param name="obj">Output EntityBase object</param>
        '''<param name="exception">output Exception value if deserialize failed</param>
        '''<returns>true if this XmlSerializer can deserialize the object; otherwise, false</returns>
        Public Overloads Shared Function Deserialize(ByVal xml As String, ByRef obj As T, ByRef exception As System.Exception) As Boolean
            exception = Nothing
            obj = CType(Nothing, T)
            Try 
                obj = Deserialize(xml)
                Return true
            Catch ex As System.Exception
                exception = ex
                Return false
            End Try
        End Function
        
        Public Overloads Shared Function Deserialize(ByVal xml As String, ByRef obj As T) As Boolean
            Dim exception As System.Exception = Nothing
            Return Deserialize(xml, obj, exception)
        End Function
        
        Public Overloads Shared Function Deserialize(ByVal xml As String) As T
            Dim stringReader As System.IO.StringReader = Nothing
            Try 
                stringReader = New System.IO.StringReader(xml)
                Return CType(Serializer.Deserialize(System.Xml.XmlReader.Create(stringReader)),T)
            Finally
                If (Not (stringReader) Is Nothing) Then
                    stringReader.Dispose
                End If
            End Try
        End Function
        
        '''<summary>
        '''Serializes current EntityBase object into file
        '''</summary>
        '''<param name="fileName">full path of outupt xml file</param>
        '''<param name="exception">output Exception value if failed</param>
        '''<returns>true if can serialize and save into file; otherwise, false</returns>
        Public Overloads Overridable Function SaveToFile(ByVal fileName As String, ByRef exception As System.Exception) As Boolean
            exception = Nothing
            Try 
                SaveToFile(fileName)
                Return true
            Catch e As System.Exception
                exception = e
                Return false
            End Try
        End Function
        
        Public Overloads Overridable Sub SaveToFile(ByVal fileName As String)
            Dim streamWriter As System.IO.StreamWriter = Nothing
            Try 
                Dim xmlString As String = Serialize
                Dim xmlFile As System.IO.FileInfo = New System.IO.FileInfo(fileName)
                streamWriter = xmlFile.CreateText
                streamWriter.WriteLine(xmlString)
                streamWriter.Close
            Finally
                If (Not (streamWriter) Is Nothing) Then
                    streamWriter.Dispose
                End If
            End Try
        End Sub
        
        '''<summary>
        '''Deserializes xml markup from file into an EntityBase object
        '''</summary>
        '''<param name="fileName">string xml file to load and deserialize</param>
        '''<param name="obj">Output EntityBase object</param>
        '''<param name="exception">output Exception value if deserialize failed</param>
        '''<returns>true if this XmlSerializer can deserialize the object; otherwise, false</returns>
        Public Overloads Shared Function LoadFromFile(ByVal fileName As String, ByRef obj As T, ByRef exception As System.Exception) As Boolean
            exception = Nothing
            obj = CType(Nothing, T)
            Try 
                obj = LoadFromFile(fileName)
                Return true
            Catch ex As System.Exception
                exception = ex
                Return false
            End Try
        End Function
        
        Public Overloads Shared Function LoadFromFile(ByVal fileName As String, ByRef obj As T) As Boolean
            Dim exception As System.Exception = Nothing
            Return LoadFromFile(fileName, obj, exception)
        End Function
        
        Public Overloads Shared Function LoadFromFile(ByVal fileName As String) As T
            Dim file As System.IO.FileStream = Nothing
            Dim sr As System.IO.StreamReader = Nothing
            Try 
                file = New System.IO.FileStream(fileName, FileMode.Open, FileAccess.Read)
                sr = New System.IO.StreamReader(file)
                Dim xmlString As String = sr.ReadToEnd
                sr.Close
                file.Close
                Return Deserialize(xmlString)
            Finally
                If (Not (file) Is Nothing) Then
                    file.Dispose
                End If
                If (Not (sr) Is Nothing) Then
                    sr.Dispose
                End If
            End Try
        End Function
        #End Region
    End Class
    #End Region
    
    <System.CodeDom.Compiler.GeneratedCodeAttribute("Xsd2Code", "3.4.0.32990"),  _
     System.SerializableAttribute(),  _
     System.ComponentModel.DesignerCategoryAttribute("code"),  _
     System.Xml.Serialization.XmlTypeAttribute([Namespace]:="http://tempuri.org/Module.xs"),  _
     System.Xml.Serialization.XmlRootAttribute("StandardModule", [Namespace]:="http://tempuri.org/Module.xs", IsNullable:=false)>  _
    Partial Public Class ModuleType
        Inherits EntityBase(Of ModuleType)
        
        Private elementsField As List(Of FModuleElement)
        
        Private filtersField As List(Of FilterType)
        
        Private internalHardwareField As ModuleTypeInternalHardware
        
        Private preFilterClampField As List(Of MiscSystemType)
        
        Private intermediateFilterClampField As List(Of MiscSystemType)
        
        Private guardFilterClampField As List(Of MiscSystemType)
        
        Private liftingLugsField As List(Of MiscSystemType)
        
        Private laddersField As List(Of MiscSystemType)
        
        Private safetyGatesField As List(Of MiscSystemType)
        
        Private assemblyFilenameField As String
        
        Private moduleNumFiltersHighField As String
        
        Private moduleNumFiltersWideField As String
        
        Private includesVerticalSplitField As Boolean
        
        Private includesHorizontalSplitField As Boolean
        
        Private drainPanRequiredField As Boolean
        
        Private isUpperModuleField As Boolean
        
        Private isLowerModuleField As Boolean
        
        Public Sub New()
            MyBase.New
            Me.safetyGatesField = New List(Of MiscSystemType)
            Me.laddersField = New List(Of MiscSystemType)
            Me.liftingLugsField = New List(Of MiscSystemType)
            Me.guardFilterClampField = New List(Of MiscSystemType)
            Me.intermediateFilterClampField = New List(Of MiscSystemType)
            Me.preFilterClampField = New List(Of MiscSystemType)
            Me.internalHardwareField = New ModuleTypeInternalHardware
            Me.filtersField = New List(Of FilterType)
            Me.elementsField = New List(Of FModuleElement)
            Me.assemblyFilenameField = "Something.iam"
            Me.moduleNumFiltersHighField = "4"
            Me.moduleNumFiltersWideField = "7"
            Me.includesVerticalSplitField = false
            Me.includesHorizontalSplitField = false
            Me.drainPanRequiredField = false
            Me.isUpperModuleField = false
            Me.isLowerModuleField = false
        End Sub
        
        <System.Xml.Serialization.XmlArrayItemAttribute("Element", IsNullable:=false)>  _
        Public Property Elements() As List(Of FModuleElement)
            Get
                Return Me.elementsField
            End Get
            Set
                Me.elementsField = value
            End Set
        End Property
        
        <System.Xml.Serialization.XmlArrayItemAttribute("Filter", IsNullable:=false)>  _
        Public Property Filters() As List(Of FilterType)
            Get
                Return Me.filtersField
            End Get
            Set
                Me.filtersField = value
            End Set
        End Property
        
        Public Property InternalHardware() As ModuleTypeInternalHardware
            Get
                Return Me.internalHardwareField
            End Get
            Set
                Me.internalHardwareField = value
            End Set
        End Property
        
        <System.Xml.Serialization.XmlElementAttribute("PreFilterClamp")>  _
        Public Property PreFilterClamp() As List(Of MiscSystemType)
            Get
                Return Me.preFilterClampField
            End Get
            Set
                Me.preFilterClampField = value
            End Set
        End Property
        
        <System.Xml.Serialization.XmlElementAttribute("IntermediateFilterClamp")>  _
        Public Property IntermediateFilterClamp() As List(Of MiscSystemType)
            Get
                Return Me.intermediateFilterClampField
            End Get
            Set
                Me.intermediateFilterClampField = value
            End Set
        End Property
        
        <System.Xml.Serialization.XmlElementAttribute("GuardFilterClamp")>  _
        Public Property GuardFilterClamp() As List(Of MiscSystemType)
            Get
                Return Me.guardFilterClampField
            End Get
            Set
                Me.guardFilterClampField = value
            End Set
        End Property
        
        <System.Xml.Serialization.XmlElementAttribute("LiftingLugs")>  _
        Public Property LiftingLugs() As List(Of MiscSystemType)
            Get
                Return Me.liftingLugsField
            End Get
            Set
                Me.liftingLugsField = value
            End Set
        End Property
        
        <System.Xml.Serialization.XmlElementAttribute("Ladders")>  _
        Public Property Ladders() As List(Of MiscSystemType)
            Get
                Return Me.laddersField
            End Get
            Set
                Me.laddersField = value
            End Set
        End Property
        
        <System.Xml.Serialization.XmlElementAttribute("SafetyGates")>  _
        Public Property SafetyGates() As List(Of MiscSystemType)
            Get
                Return Me.safetyGatesField
            End Get
            Set
                Me.safetyGatesField = value
            End Set
        End Property
        
        <System.Xml.Serialization.XmlAttributeAttribute(),  _
         System.ComponentModel.DefaultValueAttribute("Something.iam")>  _
        Public Property AssemblyFilename() As String
            Get
                Return Me.assemblyFilenameField
            End Get
            Set
                Me.assemblyFilenameField = value
            End Set
        End Property
        
        <System.Xml.Serialization.XmlAttributeAttribute(DataType:="integer"),  _
         System.ComponentModel.DefaultValueAttribute("4")>  _
        Public Property ModuleNumFiltersHigh() As String
            Get
                Return Me.moduleNumFiltersHighField
            End Get
            Set
                Me.moduleNumFiltersHighField = value
            End Set
        End Property
        
        <System.Xml.Serialization.XmlAttributeAttribute(DataType:="integer"),  _
         System.ComponentModel.DefaultValueAttribute("7")>  _
        Public Property ModuleNumFiltersWide() As String
            Get
                Return Me.moduleNumFiltersWideField
            End Get
            Set
                Me.moduleNumFiltersWideField = value
            End Set
        End Property
        
        <System.Xml.Serialization.XmlAttributeAttribute(),  _
         System.ComponentModel.DefaultValueAttribute(false)>  _
        Public Property IncludesVerticalSplit() As Boolean
            Get
                Return Me.includesVerticalSplitField
            End Get
            Set
                Me.includesVerticalSplitField = value
            End Set
        End Property
        
        <System.Xml.Serialization.XmlAttributeAttribute(),  _
         System.ComponentModel.DefaultValueAttribute(false)>  _
        Public Property IncludesHorizontalSplit() As Boolean
            Get
                Return Me.includesHorizontalSplitField
            End Get
            Set
                Me.includesHorizontalSplitField = value
            End Set
        End Property
        
        <System.Xml.Serialization.XmlAttributeAttribute(),  _
         System.ComponentModel.DefaultValueAttribute(false)>  _
        Public Property DrainPanRequired() As Boolean
            Get
                Return Me.drainPanRequiredField
            End Get
            Set
                Me.drainPanRequiredField = value
            End Set
        End Property
        
        <System.Xml.Serialization.XmlAttributeAttribute(),  _
         System.ComponentModel.DefaultValueAttribute(false)>  _
        Public Property IsUpperModule() As Boolean
            Get
                Return Me.isUpperModuleField
            End Get
            Set
                Me.isUpperModuleField = value
            End Set
        End Property
        
        <System.Xml.Serialization.XmlAttributeAttribute(),  _
         System.ComponentModel.DefaultValueAttribute(false)>  _
        Public Property IsLowerModule() As Boolean
            Get
                Return Me.isLowerModuleField
            End Get
            Set
                Me.isLowerModuleField = value
            End Set
        End Property
    End Class
    
    <System.CodeDom.Compiler.GeneratedCodeAttribute("Xsd2Code", "3.4.0.32990"),  _
     System.SerializableAttribute(),  _
     System.ComponentModel.DesignerCategoryAttribute("code"),  _
     System.Xml.Serialization.XmlTypeAttribute([Namespace]:="http://tempuri.org/Module.xs"),  _
     System.Xml.Serialization.XmlRootAttribute([Namespace]:="http://tempuri.org/Module.xs", IsNullable:=true)>  _
    Partial Public Class FModuleElement
        Inherits EntityBase(Of FModuleElement)
        
        Private assemblyFilenameField As String
        
        Private elementIDField As String
        
        Private parentModuleIDField As String
        
        Private elementIsDisabledField As Boolean
        
        Private elementBracingLeftFrontField As Boolean
        
        Private elementBracingLeftFrontFieldSpecified As Boolean
        
        Private elementBracingLeftRearField As Boolean
        
        Private elementBracingLeftRearFieldSpecified As Boolean
        
        Private elementBracingRightFrontField As Boolean
        
        Private elementBracingRightFrontFieldSpecified As Boolean
        
        Private elementBracingRightRearField As Boolean
        
        Private elementBracingRightRearFieldSpecified As Boolean
        
        Private elementWidthField As String
        
        Public Sub New()
            MyBase.New
            Me.assemblyFilenameField = "Something.iam"
            Me.elementIsDisabledField = false
        End Sub
        
        <System.Xml.Serialization.XmlAttributeAttribute(),  _
         System.ComponentModel.DefaultValueAttribute("Something.iam")>  _
        Public Property AssemblyFilename() As String
            Get
                Return Me.assemblyFilenameField
            End Get
            Set
                Me.assemblyFilenameField = value
            End Set
        End Property
        
        <System.Xml.Serialization.XmlAttributeAttribute()>  _
        Public Property ElementID() As String
            Get
                Return Me.elementIDField
            End Get
            Set
                Me.elementIDField = value
            End Set
        End Property
        
        <System.Xml.Serialization.XmlAttributeAttribute()>  _
        Public Property ParentModuleID() As String
            Get
                Return Me.parentModuleIDField
            End Get
            Set
                Me.parentModuleIDField = value
            End Set
        End Property
        
        <System.Xml.Serialization.XmlAttributeAttribute(),  _
         System.ComponentModel.DefaultValueAttribute(false)>  _
        Public Property ElementIsDisabled() As Boolean
            Get
                Return Me.elementIsDisabledField
            End Get
            Set
                Me.elementIsDisabledField = value
            End Set
        End Property
        
        <System.Xml.Serialization.XmlAttributeAttribute()>  _
        Public Property ElementBracingLeftFront() As Boolean
            Get
                Return Me.elementBracingLeftFrontField
            End Get
            Set
                Me.elementBracingLeftFrontField = value
            End Set
        End Property
        
        <System.Xml.Serialization.XmlIgnoreAttribute()>  _
        Public Property ElementBracingLeftFrontSpecified() As Boolean
            Get
                Return Me.elementBracingLeftFrontFieldSpecified
            End Get
            Set
                Me.elementBracingLeftFrontFieldSpecified = value
            End Set
        End Property
        
        <System.Xml.Serialization.XmlAttributeAttribute()>  _
        Public Property ElementBracingLeftRear() As Boolean
            Get
                Return Me.elementBracingLeftRearField
            End Get
            Set
                Me.elementBracingLeftRearField = value
            End Set
        End Property
        
        <System.Xml.Serialization.XmlIgnoreAttribute()>  _
        Public Property ElementBracingLeftRearSpecified() As Boolean
            Get
                Return Me.elementBracingLeftRearFieldSpecified
            End Get
            Set
                Me.elementBracingLeftRearFieldSpecified = value
            End Set
        End Property
        
        <System.Xml.Serialization.XmlAttributeAttribute()>  _
        Public Property ElementBracingRightFront() As Boolean
            Get
                Return Me.elementBracingRightFrontField
            End Get
            Set
                Me.elementBracingRightFrontField = value
            End Set
        End Property
        
        <System.Xml.Serialization.XmlIgnoreAttribute()>  _
        Public Property ElementBracingRightFrontSpecified() As Boolean
            Get
                Return Me.elementBracingRightFrontFieldSpecified
            End Get
            Set
                Me.elementBracingRightFrontFieldSpecified = value
            End Set
        End Property
        
        <System.Xml.Serialization.XmlAttributeAttribute()>  _
        Public Property ElementBracingRightRear() As Boolean
            Get
                Return Me.elementBracingRightRearField
            End Get
            Set
                Me.elementBracingRightRearField = value
            End Set
        End Property
        
        <System.Xml.Serialization.XmlIgnoreAttribute()>  _
        Public Property ElementBracingRightRearSpecified() As Boolean
            Get
                Return Me.elementBracingRightRearFieldSpecified
            End Get
            Set
                Me.elementBracingRightRearFieldSpecified = value
            End Set
        End Property
        
        <System.Xml.Serialization.XmlAttributeAttribute(DataType:="integer")>  _
        Public Property ElementWidth() As String
            Get
                Return Me.elementWidthField
            End Get
            Set
                Me.elementWidthField = value
            End Set
        End Property
    End Class
    
    <System.CodeDom.Compiler.GeneratedCodeAttribute("Xsd2Code", "3.4.0.32990"),  _
     System.SerializableAttribute(),  _
     System.ComponentModel.DesignerCategoryAttribute("code"),  _
     System.Xml.Serialization.XmlTypeAttribute([Namespace]:="http://tempuri.org/Module.xs"),  _
     System.Xml.Serialization.XmlRootAttribute([Namespace]:="http://tempuri.org/Module.xs", IsNullable:=true)>  _
    Partial Public Class MiscSystemType
        Inherits EntityBase(Of MiscSystemType)
        
        Private assemblyFilenameField As String
        
        Private systemTypeField As String
        
        Private systemVersionField As String
        
        Private textField As List(Of String)
        
        Public Sub New()
            MyBase.New
            Me.textField = New List(Of String)
            Me.assemblyFilenameField = "Something.iam"
        End Sub
        
        <System.Xml.Serialization.XmlAttributeAttribute(),  _
         System.ComponentModel.DefaultValueAttribute("Something.iam")>  _
        Public Property AssemblyFilename() As String
            Get
                Return Me.assemblyFilenameField
            End Get
            Set
                Me.assemblyFilenameField = value
            End Set
        End Property
        
        <System.Xml.Serialization.XmlAttributeAttribute()>  _
        Public Property SystemType() As String
            Get
                Return Me.systemTypeField
            End Get
            Set
                Me.systemTypeField = value
            End Set
        End Property
        
        <System.Xml.Serialization.XmlAttributeAttribute()>  _
        Public Property SystemVersion() As String
            Get
                Return Me.systemVersionField
            End Get
            Set
                Me.systemVersionField = value
            End Set
        End Property
        
        <System.Xml.Serialization.XmlTextAttribute()>  _
        Public Property Text() As List(Of String)
            Get
                Return Me.textField
            End Get
            Set
                Me.textField = value
            End Set
        End Property
    End Class
    
    <System.CodeDom.Compiler.GeneratedCodeAttribute("Xsd2Code", "3.4.0.32990"),  _
     System.SerializableAttribute(),  _
     System.ComponentModel.DesignerCategoryAttribute("code"),  _
     System.Xml.Serialization.XmlTypeAttribute([Namespace]:="http://tempuri.org/Module.xs"),  _
     System.Xml.Serialization.XmlRootAttribute([Namespace]:="http://tempuri.org/Module.xs", IsNullable:=true)>  _
    Partial Public Class FilterType
        Inherits EntityBase(Of FilterType)
        
        Private assemblyFilenameField As String
        
        Private brandField As String
        
        Private filterWidthField As Double
        
        Private filterHeightField As Double
        
        Private filterLengthField As Double
        
        Public Sub New()
            MyBase.New
            Me.assemblyFilenameField = "Something.iam"
            Me.brandField = "CompanyName"
            Me.filterWidthField = 610
            Me.filterHeightField = 610
            Me.filterLengthField = 1337
        End Sub
        
        <System.Xml.Serialization.XmlAttributeAttribute(),  _
         System.ComponentModel.DefaultValueAttribute("Something.iam")>  _
        Public Property AssemblyFilename() As String
            Get
                Return Me.assemblyFilenameField
            End Get
            Set
                Me.assemblyFilenameField = value
            End Set
        End Property
        
        <System.Xml.Serialization.XmlAttributeAttribute(),  _
         System.ComponentModel.DefaultValueAttribute("CompanyName")>  _
        Public Property Brand() As String
            Get
                Return Me.brandField
            End Get
            Set
                Me.brandField = value
            End Set
        End Property
        
        <System.Xml.Serialization.XmlAttributeAttribute(),  _
         System.ComponentModel.DefaultValueAttribute(610)>  _
        Public Property FilterWidth() As Double
            Get
                Return Me.filterWidthField
            End Get
            Set
                Me.filterWidthField = value
            End Set
        End Property
        
        <System.Xml.Serialization.XmlAttributeAttribute(),  _
         System.ComponentModel.DefaultValueAttribute(610)>  _
        Public Property FilterHeight() As Double
            Get
                Return Me.filterHeightField
            End Get
            Set
                Me.filterHeightField = value
            End Set
        End Property
        
        <System.Xml.Serialization.XmlAttributeAttribute(),  _
         System.ComponentModel.DefaultValueAttribute(1337)>  _
        Public Property FilterLength() As Double
            Get
                Return Me.filterLengthField
            End Get
            Set
                Me.filterLengthField = value
            End Set
        End Property
    End Class
    
    <System.CodeDom.Compiler.GeneratedCodeAttribute("Xsd2Code", "3.4.0.32990"),  _
     System.SerializableAttribute(),  _
     System.ComponentModel.DesignerCategoryAttribute("code"),  _
     System.Xml.Serialization.XmlTypeAttribute(AnonymousType:=true, [Namespace]:="http://tempuri.org/Module.xs")>  _
    Partial Public Class ModuleTypeInternalHardware
        Inherits EntityBase(Of ModuleTypeInternalHardware)
        
        Private gratingField As List(Of MiscSystemType)
        
        Private dripPanField As List(Of MiscSystemType)
        
        Private gasketField As List(Of MiscSystemType)
        
        '''<summary>
        '''ModuleTypeInternalHardware class constructor
        '''</summary>
        Public Sub New()
            MyBase.New
            Me.gasketField = New List(Of MiscSystemType)
            Me.dripPanField = New List(Of MiscSystemType)
            Me.gratingField = New List(Of MiscSystemType)
        End Sub
        
        <System.Xml.Serialization.XmlElementAttribute("Grating")>  _
        Public Property Grating() As List(Of MiscSystemType)
            Get
                Return Me.gratingField
            End Get
            Set
                Me.gratingField = value
            End Set
        End Property
        
        <System.Xml.Serialization.XmlElementAttribute("DripPan")>  _
        Public Property DripPan() As List(Of MiscSystemType)
            Get
                Return Me.dripPanField
            End Get
            Set
                Me.dripPanField = value
            End Set
        End Property
        
        <System.Xml.Serialization.XmlElementAttribute("Gasket")>  _
        Public Property Gasket() As List(Of MiscSystemType)
            Get
                Return Me.gasketField
            End Get
            Set
                Me.gasketField = value
            End Set
        End Property
    End Class
End Namespace
