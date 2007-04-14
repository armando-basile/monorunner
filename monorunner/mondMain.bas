Attribute VB_Name = "modMain"
Option Explicit

Sub Main()
    'Entry Point dell'applicazione
    Dim j As Integer
    Dim monoPath As String
    Dim monoRepository As String
    Dim monoKH As Long
    Dim myKey As Long
    Dim myKeys As Variant
    Dim numKeys As Long
    Dim monoApplication As String
    
    monoApplication = Trim(Command$)
    If monoApplication = "" Then End
    
    'Find mono path
    If RegOpenKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Novell\Mono", monoKH) = False Then
        MsgBox "NO MONO FRAMEWORK FOUNDED... download it from http://www.mono-project.com", _
            vbCritical, _
            App.Title & " v" & App.Major & "." & App.Minor & App.Revision
        End
    End If
    myKeys = RegEnumKey(monoKH)
    
    If numKeys = Null Then
        MsgBox "NO MONO FRAMEWORK FOUNDED... download it from http://www.mono-project.com", _
            vbCritical, _
            App.Title & " v" & App.Major & "." & App.Minor & App.Revision
        End
    End If
    
    For j = 0 To UBound(myKeys)
        Debug.Print myKeys(j)
        If monoRepository < myKeys(j) Then monoRepository = myKeys(j)
    Next j
    
    monoPath = GetRegistryValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Novell\Mono\" & monoRepository, "SdkInstallRoot")
    If monoPath = "" Then
        MsgBox "NO MONO FRAMEWORK FOUNDED... download it from http://www.mono-project.com", _
            vbCritical, _
            App.Title & " v" & App.Major & "." & App.Minor & App.Revision
        End
    End If
    Debug.Print monoPath
    
    Shell monoPath & "\bin\mono.exe """ & monoApplication & """", vbNormalFocus
    
    
End Sub
