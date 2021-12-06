' Simple VBScript to Export configuration manager Platform/Configuration mismatches
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

csvFile.WriteLine("Project" & CsvSep & "Solution Conf" & CsvSep & "Project Conf" _
        & CsvSep & "Solution Platform" & CsvSep & "Project Platform" & CsvSep & "Build" & CsvSep & "Mismatch")

Dim sc
For Each sc in sb.SolutionConfigurations
	' PlatformName is
	' on https://docs.microsoft.com/en-us/dotnet/api/envdte80.solutionconfiguration2?view=visualstudiosdk-2022
	WScript.Echo("SC: '" & sc.Name & "|" & sc.PlatformName )
	Dim ssc
	For Each ssc in sc.SolutionContexts
	    'WScript.Echo("    " & ssc.ProjectName & " " & ssc.ConfigurationName & "|" & ssc.PlatformName & " ShouldBuild? " & ssc.ShouldBuild)
	    Dim Mismatch
	    Mismatch = False
	    If sc.Name <> ssc.ConfigurationName or sc.PlatformName <> ssc.PlatformName Then
		Mismatch = True
	    End If
	    csvFile.WriteLine( ssc.ProjectName & CsvSep  & sc.Name & CsvSep & ssc.ConfigurationName _
	                  & CsvSep & sc.PlatformName & CsvSep & ssc.PlatformName &  CsvSep & ssc.ShouldBuild _
			  & CsvSep & Mismatch )
	Next
Next

csvFile.Close
