Attribute VB_Name = "modRegister"
Option Explicit

Private lResult As Long

Private Const KEY_ALL_ACCESS As Long = &H2003F
'Tipi di primo livello per le chiavi del registro di configurazione
Public Const HKEY_CLASSES_ROOT As Long = &H80000000
Public Const HKEY_CURRENT_USER As Long = &H80000001
Public Const HKEY_LOCAL_MACHINE As Long = &H80000002
Public Const HKEY_USERS As Long = &H80000003
Private Const ERROR_SUCCESS As Long = 0
Private Const REG_SZ As Long = 1        'Stringa Unicode con terminazione Null
Private Const REG_BINARY As Long = 3
Private Const REG_DWORD As Long = 4     'Numero a 32 bit

Private Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type

'"RegEnumValue" restituisce Null oppure un array (0 to 2).
'Qui sono elencati gli indici con i rispettivi significati.
Private Enum enmDirValueIndex
    erValueName
    erValue
    erValueType
End Enum

Private Enum enmDataType
    erSTRING
    erByte
    erDWord
End Enum

Private Declare Function OSRegCreateKey Lib "Advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpszSubKey As String, phkResult As Long) As Long
Private Declare Function OSRegDeleteKey Lib "Advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpszSubKey As String) As Long
Private Declare Function OSRegDeleteValue Lib "Advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function OSRegOpenKeyEx Lib "Advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function OSRegQueryValueEx Lib "Advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function OSRegSetValueEx Lib "Advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpszValueName As String, ByVal dwReserved As Long, ByVal fdwType As Long, lpbData As Any, ByVal cbData As Long) As Long
Private Declare Function OSRegCloseKey Lib "Advapi32.dll" Alias "RegCloseKey" (ByVal hKey As Long) As Long
Private Declare Function OSRegQueryInfoKey Lib "Advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As Any) As Long
Private Declare Function OSRegEnumKeyEx Lib "Advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
Private Declare Function OSRegEnumValue Lib "Advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long





Public Function RegEnumValue(hKey As Long, ByRef NumValues As Long) As Variant
'------------------
'Valori in entrata:
'hKey = l'handle di una chiave aperta.
'------------------
'Valori in uscita:
'Se la funzione ha esito corretto, GetEnumValue contiene un array di tipo Variant contenente tutte i valori della chiave specificata.
'"NumValues" conterrà il numero di valori trovati.
'In caso contrario, GetEnumValue sarà uguale a Null e "NumValue" sarà 0.
'------------------
    
    Dim lngValueCount As Long, lngValueNameLength As Long, lngValueLength As Long, lngValueType As Long
    Dim strKeyName As String
    Dim lngKeyLenght As Long, lngI As Long, lngJ As Long, lngK As Long
    Dim strValueName As String
    Dim abytValue(1024) As Byte
    Dim astrValue() As String
    Dim lngTemp As Long
    Dim dblTemp As Double
    Dim strTemp As String
    Dim lngSubKeysCount As Long, lngMaxSubKeyLen As Long, lngMaxClassLen As Long, lngValues As Long, lngMaxValueNameLen As Long, lngMaxValueLen As Long, lngSecurityDescriptor As Long
    Dim udtData As FILETIME
    
    lngKeyLenght = 1024
    strKeyName = String$(1024, 32)
    lngValueCount = 0
    With udtData
        .dwHighDateTime = 0
        .dwLowDateTime = 0
    End With
    'Recupera informazioni sulla chiave.
    Call OSRegQueryInfoKey(hKey, strKeyName, lngKeyLenght, ByVal 0&, lngSubKeysCount, lngMaxSubKeyLen, lngMaxClassLen, lngValueCount, lngMaxValueNameLen, lngMaxValueLen, lngSecurityDescriptor, udtData)
    If lngValueCount > 0 Then
        NumValues = lngValueCount
        For lngI = 0 To (lngValueCount - 1)
            lngValueNameLength = 1024
            strValueName = String$(1024, 32)
            lngValueLength = 1024
            abytValue(1024) = 0
            'Recupera i nomi dei valori della chiave.
            Call OSRegEnumValue(hKey, lngI, strValueName, lngValueNameLength, ByVal 0&, lngValueType, abytValue(0), lngValueLength)
            'Ridimensiona l'array.
            ReDim Preserve astrValue(0 To 2, 0 To lngI) As String
            astrValue(erValueName, lngI) = Left(strValueName, lngValueNameLength)
            If lngValueLength > 0 Then
                Select Case lngValueType
                    Case REG_BINARY
                        lngTemp = 0
                        astrValue(erValueType, lngI) = erByte
                        For lngJ = 0 To lngValueLength - 1
                            astrValue(erValue, lngI) = astrValue(erValue, lngI) & Hex(abytValue(lngJ)) & " "
                        Next lngJ
                        astrValue(erValue, lngI) = Trim(astrValue(erValue, lngI))
                    Case REG_DWORD
                        dblTemp = 0
                        astrValue(erValueType, lngI) = erDWord
                        For lngJ = 0 To lngValueLength - 1 Step 2
                            dblTemp = dblTemp + (256& ^ lngJ) * abytValue(lngJ)
                            dblTemp = dblTemp + (256& ^ (lngJ + 1)) * abytValue(lngJ + 1)
                            strTemp = Hex(256& * abytValue(lngValueLength - 1 - lngJ) + abytValue(lngValueLength - 1 - lngJ - 1))
                            astrValue(erValue, lngI) = astrValue(erValue, lngI) & String(4 - Len(strTemp), "0") & strTemp
                        Next lngJ
                        astrValue(erValue, lngI) = astrValue(erValue, lngI) & " (" & Format(dblTemp, "0,000,000,000") & ")"
                    Case Else
                        astrValue(erValueType, lngI) = erSTRING
                        For lngJ = 0 To lngValueLength - 2
                            astrValue(erValue, lngI) = astrValue(erValue, lngI) & Chr(abytValue(lngJ))
                        Next lngJ
                End Select
            End If
        Next lngI
        RegEnumValue = astrValue
    Else
        NumValues = 0
        RegEnumValue = Null
    End If
End Function

Public Function RegEnumKey(hKey As Long) As Variant
'------------------
'Valori in entrata:
'hKey = l'handle di una chiave aperta.
'------------------
'Valori in uscita:
'Se la funzione ha esito corretto, GetEnumKey contiene un array di tipo Variant contente tutte le sottochiavi della chiave specificata.
'In caso contrario, GetEnumKey sarà uguale a Null.
'------------------
    
    Dim lngSubKeysCount As Long, lngKeyLenght As Long, L As Long
    Dim strKeyName As String, astrKey() As String
    Dim lngMaxSubKeyLen As Long, lngMaxClassLen As Long, lngValues As Long, lngMaxValueNameLen As Long, lngMaxValueLen As Long, lngSecurityDescriptor As Long
    Dim strClass As String
    Dim lngClass As Long
    Dim udtData As FILETIME

    With udtData
        .dwHighDateTime = 0
        .dwLowDateTime = 0
    End With
    
    lngKeyLenght = 1024
    strKeyName = String$(1024, 32)
    'Reucpera informazioni sulla chiave.
    Call OSRegQueryInfoKey(hKey, strKeyName, lngKeyLenght, ByVal 0&, lngSubKeysCount, lngMaxSubKeyLen, lngMaxClassLen, lngValues, lngMaxValueNameLen, lngMaxValueLen, lngSecurityDescriptor, udtData)
    If lngSubKeysCount > 0 Then
        For L = 0 To (lngSubKeysCount - 1)
            strClass = String$(1024, 32)
            lngClass = 1024
            strKeyName = String$(1024, 32)
            lngKeyLenght = 1024
            With udtData
                .dwHighDateTime = 0
                .dwLowDateTime = 0
            End With
            'Recupera il nome delle sottochiavi.
            Call OSRegEnumKeyEx(hKey, L, strKeyName, lngKeyLenght, ByVal 0&, strClass, lngClass, udtData)
            'Riminesiona l'array.
            ReDim Preserve astrKey(0 To L) As String
            astrKey(L) = Left$(strKeyName, lngKeyLenght)
        Next L
        RegEnumKey = astrKey
    Else
        RegEnumKey = Null
    End If
End Function
Public Function RegCloseKey(hKey As Long) As Boolean
'------------------
'Valori in entrata:
'hKey = L'identificativo di una chiave aperta o creata.
'------------------
'Valori in uscita:
'Se la funzione ha esito corretto (True), la chiave specificata viene chiusa.
'------------------
    On Error GoTo 0
    lResult = OSRegCloseKey(hKey)
    If lResult = ERROR_SUCCESS Then
        RegCloseKey = True
    Else
        RegCloseKey = False
    End If
End Function

Public Function RegDeleteKey(hKey As Long, lpszSubKey As String) As Boolean
'------------------
'Valori in entrata:
'hKey = HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_LOCAL_MACHINE, HKEY_USER (Chiavi principali del Registro di configurazione) o l'identificativo di una chiave aperta o creata
'lpszSubKey = La sottochiave che si vuole rimuovere (attenzione: verranno rimosse anche tutte le relative sottochiavi)
'------------------
'Valori in uscita:
'Se la funzione ha esito corretto (True), la chiave specificata viene cancellata.
'------------------
    On Error GoTo 0
    lResult = OSRegDeleteKey(hKey, lpszSubKey)
    If lResult = ERROR_SUCCESS Then
        RegDeleteKey = True
    Else
        RegDeleteKey = False
    End If
End Function

Public Function RegDeleteValue(hKey As Long, lpValueName As String) As Boolean
'------------------
'Valori in entrata:
'hKey = HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_LOCAL_MACHINE, HKEY_USER (Chiavi principali del Registro di configurazione) o l'identificativo di una chiave aperta o creata.
'lpsValueName = il nome del valore che si vuole rimuovere.
'------------------
'Valori in uscita:
'Se la funzione ha esito corretto (True), il valore specificato viene cancellato.
'------------------
    On Error GoTo 0
    lResult = OSRegDeleteValue(hKey, lpValueName)
    If lResult = ERROR_SUCCESS Then
        RegDeleteValue = True
    Else
        RegDeleteValue = False
    End If
End Function

Public Function RegOpenKey(KeyRoot As Long, KeyName As String, phkResult As Long) As Boolean
'------------------
'Valori in entrata:
'KeyRoot = HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_LOCAL_MACHINE, HKEY_USER (Chiavi principali del Registro di configurazione).
'KeyName = Percorso della chiave che si vuole aprire (ad esempio "SOFTWARE\Microsoft\Windows\CurrentVersion").
'------------------
'Valori in uscita:
'Se la funzione ha esito corretto (True), phkResult contiene l'identificativo della chiave appena creata.
'In caso contrario, phkResult sarà uguale a "".
'------------------
    On Error GoTo 0
    lResult = OSRegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, phkResult)
    If lResult = ERROR_SUCCESS Then
        RegOpenKey = True
    Else
        RegOpenKey = False
    End If
End Function

Public Function RegSetNumericValue(ByVal hKey As Long, ByVal strValueName As String, ByVal TypeOfData As Long, ByVal lData As Long) As Boolean
'------------------
'Valori in entrata:
'hKey = L'identificativo di una chiave aperta o creata.
'strValueName = Il nome del valore che si vuole creare.
'TypeOfData = REG_SZ, REG_BINARY, REG_DWORD (Tipi di dati che è possibile creare).
'lData = Dati in formato numerico che si desidera inserire nel Registro.
'------------------
'Valori in uscita:
'Se la funzione ha esito corretto (True), il valore specificato è salvato nel Registro.
'In caso contrario, nessun valore sarà salvato.
'------------------
    On Error GoTo 0
    lResult = OSRegSetValueEx(hKey, strValueName, 0&, TypeOfData, lData, 4)
    If lResult = ERROR_SUCCESS Then
        RegSetNumericValue = True
    Else
        RegSetNumericValue = False
    End If
End Function
Public Function RegSetStringValue(ByVal hKey As Long, ByVal strValueName As String, ByVal strData As String) As Boolean
'------------------
'Valori in entrata:
'hKey = L'identificativo di una chiave aperta o creata.
'strValueName = Il nome del valore che si vuole creare.
'strData = Dati in formato stringa che si desidera inserire nel Registro.
'------------------
'Valori in uscita:
'Se la funzione ha esito corretto (True), il valore specificato è salvato nel Registro.
'In caso contrario, nessun valore sarà salvato.
'------------------
    On Error GoTo 0
    If hKey = 0 Then
        Exit Function
    End If
    lResult = OSRegSetValueEx(hKey, strValueName, 0&, REG_SZ, ByVal strData, LenB(StrConv(strData, vbFromUnicode)) + 1)
    If lResult = ERROR_SUCCESS Then
        RegSetStringValue = True
    Else
        RegSetStringValue = False
    End If
End Function
Public Function RegCreateKey(ByVal hKey As Long, ByVal lpszSubKeyPermanent As String, ByVal lpszSubKeyRemovable As String, phkResult As Long) As Boolean
'------------------
'Valori in entrata:
'hKey = HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_LOCAL_MACHINE, HKEY_USER (Chiavi principali del Registro di configurazione).
'lpszSubKeyPermanent = Percorso della chiave che si vuole creare (ad esempio "SOFTWARE\Microsoft\Windows\CurrentVersion").
'lpszSubKeyRemovable = Questo parametro può tranquillamente essere uguale a "".
'------------------
'Valori in uscita:
'Se la funzione ha esito corretto (True), phkResult contiene l'identificativo della chiave appena creata.
'In caso contrario, phkResult sarà uguale a "".
'------------------
    On Error GoTo 0
    Dim strSubKeyFull As String
    If lpszSubKeyPermanent = "" Then
        RegCreateKey = False 'Errore: lpszSubKeyPermanent non può essere uguale a ""
        Exit Function
    End If
    If Left$(lpszSubKeyRemovable, 1) = "\" Then
        lpszSubKeyRemovable = Mid$(lpszSubKeyRemovable, 2)
    End If
    If lpszSubKeyRemovable <> "" Then
        strSubKeyFull = lpszSubKeyPermanent & "\" & lpszSubKeyRemovable
    Else
        strSubKeyFull = lpszSubKeyPermanent
    End If
    lResult = OSRegCreateKey(hKey, strSubKeyFull, phkResult)
    If lResult = ERROR_SUCCESS Then
        RegCreateKey = True
    Else
        RegCreateKey = False
    End If
End Function
Public Function GetRegistryValue(KeyRoot As Long, KeyName As String, QueryValue As String) As Variant
'------------------
'Valori in entrata:
'KeyRoot = HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_LOCAL_MACHINE, HKEY_USER (Chiavi di primo livello del Registro di configurazione).
'KeyName = Percorso della chiave in cui si trova il valore da ottenere (ad esempio "SOFTWARE\Microsoft\Windows\CurrentVersion").
'QueryValue = Nome del valore da ottenere (ad esempio "RegisteredOwner").
'------------------
'Valori in uscita:
'Se la funzione ha esito corretto, GetRegistryValue contiene il valore richiesto.
'In caso contrario, GetReistryValue sarà uguale a "".
'------------------
    On Error GoTo 0
    Dim Value As String
    'Richiama la routine per ottenere il valore richiesto.
    If GetKeyValue(KeyRoot, KeyName, QueryValue, Value) = True Then
        GetRegistryValue = Value
    Else
        GetRegistryValue = ""
    End If
End Function

Private Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    'Questa routine è richiamata in automatico da "GetRegistryValue".
    
    Dim I As Long           'Contatore per il ciclo
    Dim rc As Long
    Dim hKey As Long        'Handle a una chiave del registro di configurazione aperta
    Dim hDepth As Long
    Dim KeyValType As Long  'Tipo di dati di una chiave del registro di configurazione
    Dim tmpVal As String    'Variabile per la memorizzazione temporanea del valore di una chiave del registro di configurazione
    Dim KeyValSize As Long  'Dimensioni della variabile per la chiave del registro di configurazione
    '------------------------------------------------------------------
    ' Apre la chiave del registro sotto KeyRoot
    '------------------------------------------------------------------
    rc = OSRegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) 'Apre la chiave del registro
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          'Gestisce gli errori...
    tmpVal = String$(1024, 0)                             'Assegna lo spazio per la variabile
    KeyValSize = 1024                                       'Definisce le dimensioni della variabile
    '---------------------------------------------------------------
    ' Recupera il valore della chiave del registro di configurazione...
    '---------------------------------------------------------------
    rc = OSRegQueryValueEx(hKey, SubKeyRef, 0, KeyValType, tmpVal, KeyValSize)    'Recupera/crea il valore della chiave
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          'Gestisce gli errori
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           'Win95 aggiunge una stringa con terminazione Null...
        tmpVal = Left(tmpVal, KeyValSize - 1)               'Trova Null, estrae dalla stringa
    Else                                                    'WinNT non aggiunge la terminazione Null alle stringhe...
        tmpVal = Left(tmpVal, KeyValSize)                   'Non trova Null, estrae solo la stringa
    End If
    '----------------------------------------------------------------
    ' Determina il tipo del valore della chiave per la conversione...
    '----------------------------------------------------------------
    Select Case KeyValType                                      'Esamina i tipi di dati...
        Case REG_SZ                                             'Tipo di dati String per la chiave del registro
            KeyVal = tmpVal                                     'Copia il valore String
        Case REG_DWORD                                          'Tipo di dati Double Word per la chiave del registro
            For I = Len(tmpVal) To 1 Step -1                    'Converte ogni bit
                KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, I, 1)))   'Crea il valore carattere per carattere.
            Next
            KeyVal = Format$("&h" + KeyVal)                     'Converte Double Word in String
    End Select
    GetKeyValue = True                                      'Operazione riuscita
    OSRegCloseKey hKey                                  'Chiude la chiave del registro
    Exit Function                                           'Esce
GetKeyError:      'Svuota la variabile e chiude la chiave in seguito ad un errore.
    KeyVal = ""                                             'Imposta su una stringa vuota il valore restituito
    GetKeyValue = False                                     'Operazione non riuscita
    OSRegCloseKey hKey                                 'Chiude la chiave del registro
End Function
