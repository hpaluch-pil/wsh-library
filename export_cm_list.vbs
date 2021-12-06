' Simple VBScript to Export configuration manager Lists/Properties
' run as:
' cscript export_cm_list.vbs my_solution_wiht_cpp_projects.sln output.csv

Option Explicit 

' C++ project type
' See https://github.com/umbraco/Visual-Studio-Extension/blob/master/UmbracoStudio/VsConstants.cs for list
Const CppProjectTypeGuid = "{8BC9CEB8-8B4A-11D0-8D11-00A0C91BC942}"
Const CsProjectTypeGuid = "{FAE04EC0-301F-11D3-BF4B-00C04F79EFBC}"

' VisualStudio object:
'     12 => VS 2013
Const VsObjectName = "VisualStudio.DTE.12.0"

' CSV output separator. Typically "," or ";"
Const CsvSep = ","

' Simple map of toolset versions (incomplete)
' Map from: https://marcofoco.com/microsoft-visual-c-version-map/
Dim ToolsetDict
Set ToolsetDict = CreateObject("Scripting.Dictionary")
Call ToolsetDict.Add("v80","Visual Studio 2005")
Call ToolsetDict.Add("v90","Visual Studio 2008")
Call ToolsetDict.Add("v100","Visual Studio 2010")
Call ToolsetDict.Add("v110","Visual Studio 2012")
Call ToolsetDict.Add("v120","Visual Studio 2013")
Call ToolsetDict.Add("v140","Visual Studio 2015")
Call ToolsetDict.Add("v141","Visual Studio 2017")
Call ToolsetDict.Add("v142","Visual Studio 2019")


Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

Sub Odpad
	Dim ci
	For each ci in cm.ConfigurationRowNames
		WScript.Echo("CI: " & ci )
	Next

	Dim i
	For i = 1 To cm.Count
		WScript.Echo("I " & i &  ": " & cm.Item(i).ConfigurationName & "|" & cm.Item(i).PlatformName & " Build? " & cm.Item(i).IsBuildable)
		Dim pp
		Set pp = cm.Item(i).Owner
		'WScript.Echo("Parent proj: " & pp.Name)
	Next
End Sub

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

        Dim ToolsetName
        ToolsetName="N/A"

        If ToolsetDict.Exists(toolset) Then
            ToolsetName = ToolsetDict.Item(toolset)
        End If

        ' CSV output
        csvFile.WriteLine(proj.Name & CsvSep & cpCombo(0) & CsvSep & cpCombo(1) _
                 & CsvSep & toolset & CsvSep & ToolsetName )
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
'Set objDTE = CreateObject(VsObjectName)

' connect to existing VS xxxx
Set objDTE = GetObject(,VsObjectName)
 
Wscript.Echo("VS Version: " & objDTE.Version)

' Make it visible and keep it open after we finish this script
objDTE.MainWindow.Visible = True
objDTE.UserControl = True

Dim sol
Set sol = objDTE.Solution
sol.Open(WScript.Arguments.Item(ArgSlnIndex))

Dim sb
Set sb = sol.SolutionBuild

Dim sc
For Each sc in sb.SolutionConfigurations
	' PlatformName is
	' on https://docs.microsoft.com/en-us/dotnet/api/envdte80.solutionconfiguration2?view=visualstudiosdk-2022
	WScript.Echo("SC: '" & sc.Name & "|" & sc.PlatformName )
	Dim ssc
	For Each ssc in sc.SolutionContexts
	    WScript.Echo("    " & ssc.ProjectName & " " & ssc.ConfigurationName & "|" & ssc.PlatformName & " ShouldBuild? " & ssc.ShouldBuild)
	Next
	WScript.Echo("  ")
Next

csvFile.WriteLine("Project" & CsvSep & "Configuration" & CsvSep & "Platform" _
        & CsvSep & "Toolset" & CsvSep & "Toolset Name")

csvFile.Close
