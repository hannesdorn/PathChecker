'Hannes Dorn <hannes@dorn.cc>
'2017-05-21
'Version 1
'GPLv3

Option Explicit

Dim oFs
Dim oShell
Dim oEnvironment
Dim aPaths
Dim aPathsInvalid

Set oFs = CreateObject( "Scripting.FileSystemObject" )
Set oShell = WScript.CreateObject( "WScript.Shell" )

WScript.Echo "System"
Set oEnvironment = oShell.Environment( "SYSTEM" )
aPaths = Split( oEnvironment( "Path" ), ";" )
Wscript.Echo Join( ValidatePaths( aPaths ), vbCrLf )

WScript.Echo "User"
Set oEnvironment = oShell.Environment( "USER" )
aPaths = Split( oEnvironment( "Path" ), ";" )
Wscript.Echo Join( ValidatePaths( aPaths ), vbCrLf )

Function ValidatePaths( aPaths )

    Dim i
    Dim aPathsInvalid

    aPathsInvalid = Array()
    For i = LBound( aPaths ) to UBound( aPaths )
        If Not oFs.FolderExists( oShell.ExpandEnvironmentStrings( aPaths( i ) ) ) Then
            Redim Preserve aPathsInvalid( UBound( aPathsInvalid ) + 1 )
            aPathsInvalid( UBound( aPathsInvalid ) ) = aPaths( i )
        End If
    Next

    ValidatePaths = aPathsInvalid

End Function
