' Simple VBScript to extract all used MSVC Toolsets from projects in solution
' run as:
' cscript export_toolset_list.vbs my_solution_wiht_cpp_projects.sln output.csv

Option Explicit 

' C++ project type
' See https://github.com/umbraco/Visual-Studio-Extension/blob/master/UmbracoStudio/VsConstants.cs for list
Const CppProjectTypeGuid = "{8BC9CEB8-8B4A-11D0-8D11-00A0C91BC942}"

' VisualStudio object:
'     12 => VS 2013
Const VsObjectName = "VisualStudio.DTE.12.0"

' CSV output separator. Typically "," or ";"
Const CsvSep = ","

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

Sub FileExistsOrDie(fileName)
    if Not fso.FileExists(fileName) Then
        WScript.Echo("Required file '" & fileName & "' does not exist")
        WScript.Quit(1)
    End If
End Sub

Sub CheckExtensionOrDie(fileName,ext)
    if StrComp(Right(fileName,Len(ext)), ext, vbTextCompare) <> 0 Then
        WScript.Echo("File name '" & fileName & "' has wrong extension '" & Right(fileName,Len(ext)) & "' <> '" & ext & "'")
        WScript.Quit(1)
    End If
End Sub

Sub DumpProjectInfo(ByRef proj, ByRef csvFile)
    ' from https://kobyk.wordpress.com/2011/11/26/modifying-a-visual-c-2010-projects-platform-toolset-programmatically-with-ivcrulepropertystorage/
    Dim c
    For Each c in proj.Object.Configurations
        'Wscript.Echo("   c name: " & c.Name)
        Dim tsRule
        Set tsRule = c.Rules.Item("ConfigurationGeneral")

        Dim toolset
        toolset = tsRule.GetUnevaluatedPropertyValue("PlatformToolset")

        'WScript.Echo("   Toolset: " & toolset)

        Dim cpCombo
        cpCombo = Split(c.Name,"|")
        If UBound(cpCombo) <> 1 Then
            WScript.Echo("Unable to split '" & c.Name & "' with '|' -  Array count  " & UBound(cpCombo)+1 & " <> 2" )
            WScript.Quit(1)
        End If
        ' CSV output
        csvFile.WriteLine(proj.Name & CsvSep & cpCombo(0) & CsvSep & cpCombo(1) & CsvSep & toolset )
    Next
End Sub

If WScript.Arguments.Count <> 2 Then
    WScript.Echo("Usage: " & WScript.ScriptName & " my_solution_file.sln output.csv")
    WScript.Quit(1)
End If

Const ArgSlnIndex = 0
Const ArgCsvNameIndex = 1

Call CheckExtensionOrDie(WScript.Arguments.Item(ArgSlnIndex),".sln")
Call CheckExtensionOrDie(WScript.Arguments.Item(ArgCsvNameIndex),".csv")

Call FileExistsOrDie(WScript.Arguments.Item(ArgSlnIndex))

Dim csvFile
Set csvFile = fso.CreateTextFile(WScript.Arguments.Item(ArgCsvNameIndex), True)

' from https://www.mztools.com/Articles/2005/MZ2005005.aspx
Dim objDTE
' Creates an instance of the Visual Studio xxxx DTE
Set objDTE = CreateObject(VsObjectName)

' connect to existing VS xxxx
'Set objDTE = GetObject(,VsObjectName)
 
'Wscript.Echo("VS Version: " & objDTE.Version)

' Make it visible and keep it open after we finish this script
objDTE.MainWindow.Visible = True
objDTE.UserControl = True

Dim sol
Set sol = objDTE.Solution
sol.Open(WScript.Arguments.Item(ArgSlnIndex))


csvFile.WriteLine("Project" & CsvSep & "Configuration" & CsvSep & "Platform" & CsvSep & "Toolset")
Dim proj
For Each proj in sol.Projects
	'WScript.Echo("Project: '" & proj.Name & "' Kind: " & proj.Kind)

    If proj.Kind = CppProjectTypeGuid Then
        Call DumpProjectInfo(proj,csvFile)
        'WScript.Echo("  Configuration: '" & proj.ConfigurationManager.ActiveConfiguration.ConfigurationName & "'")
        'WScript.Echo("  Platform: '" & proj.ConfigurationManager.ActiveConfiguration.PlatformName & "'")
    Else
        'WScript.Echo(" WARNING: Skipping unsupported project kind") 
    End If
Next
csvFile.Close
