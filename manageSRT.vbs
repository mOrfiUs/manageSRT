Option Explicit


'Tag
Const TagName = "subtítulos gracias a Creador"
'extensión de los archivos en los que se incluirá un Tag adicional
Const extToProcess = ".fzeng.srt"
Call main

Sub main()
'Exit Sub
Dim objCurrentDir, objAllFiles, fInDir, objFSO, objOutDir, szEncoding, justName, posDot, Tag, vbRes

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    Set objCurrentDir = objFSO.GetFolder(".")
    Set objAllFiles = objCurrentDir.files
    
    'crea, si no existe, un directorio OUT para escribir los archivos
    If objFSO.FolderExists(".\OUT\") = False Then Set objOutDir = objFSO.CreateFolder(".\OUT\")
    Set objOutDir = Nothing

    szEncoding = "utf-8"
    vbRes = MsgBox("Archivos origen codificados en " & vbCrLf & "Sí -> utf-8" & vbCrLf & "No -> Windows-1252" & vbCrLf & vbCrLf & "La salida siempre en Windows-1252 (ANSI)", 35, "Convertir y normalizar subtítulos")
    If vbRes = vbCancel Then Exit Sub
    If vbRes = vbNo Then szEncoding = "Windows-1252"

    For Each fInDir In objAllFiles
        'procesa todos los archivos .srt
        If InStr(1, fInDir.Name, ".srt", vbTextCompare) > 0 Then
            posDot = InStr(1, fInDir.Name, ".", vbTextCompare)
            'solo añade Tag a  los archivos extToProcess por ejemplo (".eng.srt")
            If (posDot > 0) And (InStr(1, fInDir.Name, extToProcess, vbTextCompare) > 0) Then
                justName = Left(fInDir.Name, posDot - 1)
                Tag = AddTagToSRT(UCase(objCurrentDir.Path), justName)
            End If
            Call ConvertANSIOrUTF8(fInDir.Path, ".\OUT\" & fInDir.Name, szEncoding, "Windows-1252", Tag)
            Tag = ""
        End If
    Next
    Set objAllFiles = Nothing
    Set objCurrentDir = Nothing
    Set objFSO = Nothing
End Sub

Function nSecondsToTime(ByVal nSeconds)
Dim hours, minutes, seconds, intSeconds
    If IsNull(nSeconds) = vbTrue Then Exit Function
    If IsEmpty(nSeconds) = vbTrue Then Exit Function
    If nSeconds < (10 ^ 7) Then Exit Function
    
    intSeconds = (nSeconds / (10 ^ 7)) - 2
    hours = intSeconds \ 3600
    If hours < 10 Then hours = "0" & CStr(hours)
    intSeconds = intSeconds Mod 3600
    minutes = intSeconds \ 60
    If minutes < 10 Then minutes = "0" & CStr(minutes)
    seconds = (intSeconds Mod 60)
    If seconds < 10 Then seconds = "0" & CStr(seconds)
    nSecondsToTime = hours & ":" & minutes & ":" & seconds
End Function


Function AddTagToSRT(sFolderPath, justName)
Dim objShell, objFolder, objItem, nSeconds, szText, formatTime

    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.Namespace(sFolderPath)
    
    For Each objItem In objFolder.Items
        If objItem.IsFolder = False Then
            If (InStr(1, ".MP4.MKV.AVI", Right(objItem.Name, 4), vbTextCompare) > 0) And (UCase(Left(objItem.Name, Len(justName))) = UCase(justName)) Then
                nSeconds = CStr(objItem.ExtendedProperty("Duration"))
                formatTime = nSecondsToTime(nSeconds)
                szText = formatTime & ",069" & " --> " & formatTime & ",690" & vbCrLf & TagName & vbCrLf & vbCrLf
            End If
        End If
    Next
    AddTagToSRT = szText
    Set objFolder = Nothing
    Set objShell = Nothing
End Function

Function ConvertANSIOrUTF8(FileIn, FileOut, sFrom, sTo, Tag)
Const adTypeBinary = 1
Const adTypeText = 2

Dim oFS: Set oFS = CreateObject("Scripting.FileSystemObject")
Dim oFrom: Set oFrom = CreateObject("ADODB.Stream")
Dim sFFSpec: sFFSpec = oFS.GetAbsolutePathName(FileIn)
Dim oTo: Set oTo = CreateObject("ADODB.Stream")
Dim sTFSpec: sTFSpec = oFS.GetAbsolutePathName(FileOut)
Dim szText, fRemoveSpam, findSpam, arrSpam, posLastNumerSRT, lastNumerSRT, posLastNumerSRT2

    'archivo Origen
    oFrom.Type = adTypeText: oFrom.Charset = sFrom
    oFrom.Open: oFrom.LoadFromFile sFFSpec
    'archivo Destino
    oTo.Type = adTypeText: oTo.Charset = sTo
    oTo.Open: szText = oFrom.ReadText
    'añado dos retornos al principio y al final para evitar problemas en RemoveSpam
    szText = vbCrLf & vbCrLf & szText & vbCrLf & vbCrLf

    arrSpam = Array("www", "http://", "SubXpacio", "tusseries", "HEROES TEAM", "Traducción:", "Resincro", "traducción", "traduccion", "GrupoTS", "Sous-titrage", "Adaptation", "GrupoToo", "hdtv", "subsfactory", "Traduzione", "Prerevisione", "revisione", "addic7ed", "Sincronización:", "Subtitles", "subtítulos", "subtitulos", "bluRay", "[Castilian Spanish]", "Ripeado", "re-sincro:", "sincronizado", "TheSubFactory", "Asia-Team", "Akantor", "LeapinLar", "OpenSubtitles", "Synchro", "VeRdiKT", "Subscene", "bebos123")
    For Each findSpam In arrSpam
        Do
            fRemoveSpam = RemoveSpam(szText, findSpam, FileIn)
        Loop While fRemoveSpam > 0
    Next
    'Elimino los vbCrLf sobrantes
    Do
        szText = Replace(szText, vbCrLf & vbCrLf & vbCrLf, vbCrLf & vbCrLf)
    Loop While InStr(1, szText, vbCrLf & vbCrLf & vbCrLf, vbBinaryCompare) > 0
    
    'Elimino los vbCrLf del inicio de archivo
    Do While Left(szText, 2) = vbCrLf
        szText = Mid(szText, 3)
    Loop

    If Tag <> "" Then
      posLastNumerSRT = InStrRev(szText, vbCrLf & vbCrLf, Len(szText) - 1, vbTextCompare) + 4
      posLastNumerSRT2 = InStr(posLastNumerSRT, szText, vbCrLf, vbBinaryCompare)
      If posLastNumerSRT2 > posLastNumerSRT2 Then
        lastNumerSRT = CInt(Mid(szText, posLastNumerSRT, posLastNumerSRT2 - posLastNumerSRT)) + 1
        Tag = lastNumerSRT & vbCrLf & Tag
      Else
        MsgBox "No se pudo encontrar el número del último subtíttulo", vbInformation, "Tag = " & Tag
      End If
    End If
    oTo.WriteText szText & Tag
    oTo.SaveToFile sTFSpec
    oFrom.Close: oTo.Close
    Set oFrom = Nothing: Set oTo = Nothing: Set oFS = Nothing
End Function

Function RemoveSpam(szText, szSpam, FileIn)
Dim iPos1, iPos2, iPos3
Dim szText1, szText2, szTextBorrado
'REGEX en VBS
    iPos1 = InStr(1, szText, szSpam, vbTextCompare)
    If iPos1 < 1 Then Exit Function
    iPos2 = InStr(iPos1, szText, vbCrLf & vbCrLf, vbTextCompare)
    If iPos2 < 1 Then Exit Function
    iPos3 = InStrRev(szText, vbCrLf & vbCrLf, iPos1, vbTextCompare)
    If iPos3 < 1 Then Exit Function
    'InStrRev(string1,string2[,start[,compare]])
    'Mid(string,start[,length])
    ' si es el primer subtítulo lo ignora
    If iPos3 > 2 Then szText1 = Mid(szText, 1, iPos3)
    szText2 = Mid(szText, iPos2)
    If iPos3 > 2 Then szTextBorrado = Mid(szText, iPos3, iPos2 - iPos3)
    MsgBox FileIn & vbCrLf & "Término buscado: " & UCase(szSpam) & vbCrLf & szTextBorrado, vbInformation, "subtítulo eliminado"
    szText = szText1 & vbCrLf & szText2
    RemoveSpam = 1
End Function
