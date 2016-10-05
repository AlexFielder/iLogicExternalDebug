Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Runtime.CompilerServices
Imports Inventor

Module Extensions
    <Extension>
    Public Function Value(Param As Parameter, StringValueToRemove As String) As String
        Dim StringToEdit As String = Param.Value.ToString()
        Return StringToEdit.Replace(Param.Value.ToString, StringValueToRemove)
    End Function
End Module
