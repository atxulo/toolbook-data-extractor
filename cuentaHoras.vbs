Const sFolder = "C:\DTP-E\"
Const headerOffset = 387 ' bytes que no leemos, son la cabecera
Const recordSize = 56 ' caracteres que ocupa cada registro
Const adTypeBinary = 1 
Const adTypeText = 2
Const adModeRead = 1

' Validamos los parametros obligatorios
Set colNamedArguments = WScript.Arguments.Named
If colNamedArguments.Exists("inicio") Then
  fechaInicio = colNamedArguments.Item("inicio")
Else
  fechaInicio = 00000000
End If
If colNamedArguments.Exists("nivel") Then
  nivel = colNamedArguments.Item("nivel")
Else
  nivel = -1
End If
If colNamedArguments.Exists("horas") Then
  horasCompromiso = colNamedArguments.Item("horas")
Else
  horasCompromiso = 168
End If

' Nos preparamos para leer los ficheros
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set BinaryStream = CreateObject("ADODB.Stream")
BinaryStream.Type = adTypeBinary 
BinaryStream.Open

' Leemos cada fichero
total = 0
For Each oFile In oFSO.GetFolder(sFolder + "overbook").Files
  If UCase(oFSO.GetExtensionName(oFile.Name)) = "DBF" Then
    Wscript.Echo "-------------------------------------------------------------------------"
    Wscript.Echo "|                         Fichero " & oFile.Name & "                          |"
    Wscript.Echo "-------------------------------------------------------------------------"
    fechaFichero = Mid(oFile.Name, 3, 6) & "99"
    If (fechaFichero < fechaInicio) Then
      Wscript.Echo "Ignorado por ser anterior a " & fechaInicio
    Else
      Wscript.Echo "Fecha   Nivel H.Inicio H.Fin F.Estudiadas F.Asimiladas Fichas/h T.Estudio"
      Wscript.Echo "-------------------------------------------------------------------------"
      bin = readBinary(oFile)
      numRecords = Len(bin) / recordSize
      ikasDenbora = 0
      For i = 0 to numRecords - 1
        record = Mid(bin, (recordSize * i) + 1, recordSize - 1)
        ' Leemos cada registro
        recordData = Left(record, 8)
        recordNivel = Mid(record, 16, 2)
        recordInicio = Mid(record, 18, 5)
        recordFin = Mid(record, 23, 5)
        If (recordData < fechaInicio) Then
          Wscript.Echo recordData & " " & recordNivel & "    " & recordInicio & "   " & recordFin & " ignorado por ser anterior a " & fechaInicio
        ElseIf nivel <> -1 and recordNivel <> nivel Then
          Wscript.Echo recordData & " " & recordNivel & "    " & recordInicio & "   " & recordFin & " ignorado por ser de nivel distinto a " & nivel
        Else
          recordEstudiadas = Mid(record, 28, 3)
          recodAsimiladas = Mid(record, 31, 5)
          recordFichasPorH = Mid(record, 36, 5)
          recordIkasDenbora = CInt(Mid(record, 41, 4))
          Wscript.Echo recordData & " " & recordNivel & "    " & recordInicio & "   " & recordFin & "       " & recordEstudiadas &_
            "      " & recodAsimiladas & "       " & recordFichasPorH & "    " & recordIkasDenbora & " min."
          ikasDenbora = ikasDenbora + recordIkasDenbora
        End If
      Next
      Wscript.Echo "-------------------------------------------------------------------------"
      Wscript.Echo "Subtotal -> " & ikasDenbora & " min. - " &  (ikasDenbora / 60) & " h."
      total = total + ikasDenbora
    End If
    Wscript.Echo ""
  End if
Next
Wscript.Echo "========================================"
Wscript.Echo "Total -> " & total & " min. - " &  (total / 60) & " h."
Wscript.Echo "Faltan " & horasCompromiso - (total / 60) & " h. para llegar a " & horasCompromiso & " h."

BinaryStream.Close
Set oFSO = Nothing

Function readBinary(oFile)
  Dim a
  Dim i
  Dim ts
  Set ts = oFile.OpenAsTextStream()
  a = makeArray(oFile.size - headerOffset)
  i = 0
  ' Do not replace the following block by readBinary = by ts.readAll(), it would result in broken output, because that method is not intended for binary data 
  While i < oFile.size  'Not ts.atEndOfStream
    If (i >= headerOffset) then
      a(i - headerOffset) = ts.read(1)
    Else
      ts.read(1)  'Ignoramos este byte
    End If
    i = i + 1
  Wend
  ts.close
  readBinary = Join(a,"")
End Function

Function makeArray(n) ' Small utility function
  Dim s
  s = Space(n)
  makeArray = Split(s," ")
End Function