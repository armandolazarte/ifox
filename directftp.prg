
If .F.
   ?  "Creado por Pablo Pioli"
Endif


#include IFOX.H


**************************   DIRECTFTP CLASS   *****************************

****************************************************************************
Define Class DirectFTP as iFox of iFox.PRG OlePublic

       ErrorNumber = 0
       FTPResponseCode = 0
       FTPResponseText = ""
       Response = ""

       TransferSize = 0
       TransferredBytes = 0
       EOT = .F.

       Hidden Sockets
       Hidden SocketNumber
       Hidden ServerIP

       Hidden TransferType

       Hidden TransferSocket
       Hidden FileHandle

       Hidden TransferResult

       Dimension DirFiles[1]


***
* Function eInit
***
Protected Function eInit()


* Inicializar propiedades
this.TransferType = "A"
this.Sockets = CreateObject("Sockets")

EndFunc



***
* Function Connect
*   Stablish a connection with an FTP server
***
Function Connect(cServer, cUserName, cPassword, nPort, nAdditionalWait)
Local nPass, cData, cRes


* Verificar parametros
If Type("nPort") <> "N"
   nPort = 21
Endif

If Type("nAdditionalWait") <> "N"
   nAdditionalWait = 0
Endif



* Inicializar socket
If !this.Sockets.StartOK
   this.ErrorNumber = 1
   Return(.F.)
Endif



* Construir encabezado
this.ServerIP = this.Sockets.GetIPFromName(cServer)
If Empty(this.ServerIP)
   this.ErrorNumber = 2
   Return(.F.)
Endif



* Establecer conexión
this.SocketNumber = this.Sockets.Connect(this.ServerIP, nPort)

If this.SocketNumber == 0
   this.ErrorNumber = 3
   Return(.F.)
Endif



* Esperar saludo
nPass = 1

Do While .T.
   If nPass == 1

      If !this.Sockets.WaitFor(this.SocketNumber)
         this.ErrorNumber = 4
         Return(.F.)
      Endif

      this.FTPResponseCode = Val(Left(this.Sockets.Response, 3))
      cData = SubStr(this.Sockets.Response, 5)

      Do While .T.
         Win32Delay(1)
         cRes = this.Sockets.Read(this.SocketNumber)

         If Len(cRes) <> 0
            cData = cData + cRes
         else
            Exit
         Endif
      Enddo

      this.FTPResponseText = cData

      If this.FTPResponseCode <> 220
         this.ErrorNumber = 5
         Return(.F.)
      Endif

      nPass = 2
   else

      If nAdditionalWait > 0
         this.Sockets.WaitFor(this.SocketNumber,, nAdditionalWait)
      Endif
      Exit

   Endif
Enddo



* Enviar nombre de usuario
If this.Sockets.SendReceive(this.SocketNumber, "USER " + cUserName + NEW_LINE)

   this.FTPResponseCode = Val(Left(this.Sockets.Response, 3))
   this.FTPResponseText = SubStr(this.Sockets.Response, 5)

   If (this.FTPResponseCode <> 230) .AND. (this.FTPResponseCode <> 331)
      this.ErrorNumber = 6
      Return(.F.)
   Endif

else
   this.ErrorNumber = 4
   Return(.F.)
Endif



* Enviar clave de acceso
If this.FTPResponseCode == 331
   If this.Sockets.SendReceive(this.SocketNumber, "PASS " + cPassword + NEW_LINE)

      this.FTPResponseCode = Val(Left(this.Sockets.Response, 3))
      this.FTPResponseText = SubStr(this.Sockets.Response, 5)

      If (this.FTPResponseCode <> 230) .AND. (this.FTPResponseCode <> 202)
         this.ErrorNumber = 7
         Return(.F.)
      Endif

   else
      this.ErrorNumber = 4
      Return(.F.)
   Endif
Endif


Return(.T.)



***
* Function Close()
*   Cierra la conexion establecida con Connect
***
Function Close()

If this.SocketNumber <> 0

   If this.Sockets.SendReceive(this.SocketNumber, "QUIT" + NEW_LINE)
      If this.Sockets.SocketClosed
         this.ErrorNumber = 4
         Return(.F.)
      else
         this.FTPResponseCode = Val(Left(this.Sockets.Response, 3))
         this.FTPResponseText = SubStr(this.Sockets.Response, 5)

         If this.FTPResponseCode <> 221
            this.ErrorNumber = 8
            Return(.F.)
         Endif
      Endif
   else
      this.ErrorNumber = 4
      Return(.F.)
   Endif

   this.Sockets.Close(this.SocketNumber)
Endif

Return(.T.)



***
* Function System()
*   Envia el comando SYST
***
Function System()

this.Response = ""

If this.Sockets.SendReceive(this.SocketNumber, "SYST" + NEW_LINE)
   If this.Sockets.SocketClosed
      this.ErrorNumber = 4
      Return(.F.)
   else
      this.FTPResponseCode = Val(Left(this.Sockets.Response, 3))
      this.FTPResponseText = SubStr(this.Sockets.Response, 5)

      If this.FTPResponseCode <> 215
         this.ErrorNumber = 9
         Return(.F.)
      Endif

      this.Response = SubStr(this.Sockets.Response, 5)
   Endif
else
   this.ErrorNumber = 4
   Return(.F.)
Endif

Return(.T.)



***
* Function CD(cDir)
*   Devuelve o cambia el directorio actual
***
Function CD(cDir)

this.Response = ""

If Empty(cDir)

   * Obtener el directorio activo
   If this.Sockets.SendReceive(this.SocketNumber, "PWD" + NEW_LINE)
      If this.Sockets.SocketClosed
         this.ErrorNumber = 4
         Return(.F.)
      else
         this.FTPResponseCode = Val(Left(this.Sockets.Response, 3))
         this.FTPResponseText = SubStr(this.Sockets.Response, 5)

         If this.FTPResponseCode <> 257
            this.ErrorNumber = 9
            Return(.F.)
         Endif

         this.Response = SubStr(this.Sockets.Response, 6)
         this.Response = Left(this.Response, At('"', this.Response) - 1)
      Endif
   else
      this.ErrorNumber = 4
      Return(.F.)
   Endif

else

   * Cambiar de directorio
   If this.Sockets.SendReceive(this.SocketNumber, "CWD " + cDir + NEW_LINE)
      If this.Sockets.SocketClosed
         this.ErrorNumber = 4
         Return(.F.)
      else
         this.FTPResponseCode = Val(Left(this.Sockets.Response, 3))
         this.FTPResponseText = SubStr(this.Sockets.Response, 5)

         If this.FTPResponseCode <> 250
            this.ErrorNumber = 9
            Return(.F.)
         Endif
      Endif
   else
      this.ErrorNumber = 4
      Return(.F.)
   Endif

Endif

Return(.T.)



***
* Function CDUp()
*   Cambia al directorio padre
***
Function CDUp()

* Cambiar de directorio
If this.Sockets.SendReceive(this.SocketNumber, "CDUP" + NEW_LINE)
   If this.Sockets.SocketClosed
      this.ErrorNumber = 4
      Return(.F.)
   else
      this.FTPResponseCode = Val(Left(this.Sockets.Response, 3))
      this.FTPResponseText = SubStr(this.Sockets.Response, 5)

      If this.FTPResponseCode <> 200
         this.ErrorNumber = 9
         Return(.F.)
      Endif
   Endif
else
   this.ErrorNumber = 4
   Return(.F.)
Endif

Return(.T.)



***
* Function CreateFolder(cFolder)
*   Crea una carpeta
***
Function CreateFolder(cFolder)

If this.Sockets.SendReceive(this.SocketNumber, "MKD " + cFolder + NEW_LINE)
   If this.Sockets.SocketClosed
      this.ErrorNumber = 4
      Return(.F.)
   else
      this.FTPResponseCode = Val(Left(this.Sockets.Response, 3))
      this.FTPResponseText = SubStr(this.Sockets.Response, 5)

      If this.FTPResponseCode <> 257
         this.ErrorNumber = 9
         Return(.F.)
      Endif
   Endif
else
   this.ErrorNumber = 4
   Return(.F.)
Endif

Return(.T.)



***
* Function DeleteFolder(cFolder)
*   Elimina una carpeta
***
Function DeleteFolder(cFolder)

If this.Sockets.SendReceive(this.SocketNumber, "RMD " + cFolder + NEW_LINE)
   If this.Sockets.SocketClosed
      this.ErrorNumber = 4
      Return(.F.)
   else
      this.FTPResponseCode = Val(Left(this.Sockets.Response, 3))
      this.FTPResponseText = SubStr(this.Sockets.Response, 5)

      If this.FTPResponseCode <> 250
         this.ErrorNumber = 9
         Return(.F.)
      Endif
   Endif
else
   this.ErrorNumber = 4
   Return(.F.)
Endif

Return(.T.)



***
* Function DeleteFile(cFile)
*   Elimina un archivo
***
Function DeleteFile(cFile)

If this.Sockets.SendReceive(this.SocketNumber, "DELE " + cFile + NEW_LINE)
   If this.Sockets.SocketClosed
      this.ErrorNumber = 4
      Return(.F.)
   else
      this.FTPResponseCode = Val(Left(this.Sockets.Response, 3))
      this.FTPResponseText = SubStr(this.Sockets.Response, 5)

      If this.FTPResponseCode <> 250
         this.ErrorNumber = 9
         Return(.F.)
      Endif
   Endif
else
   this.ErrorNumber = 4
   Return(.F.)
Endif

Return(.T.)



***
* Function RenameFile(cOldName, cNewName)
*   Renombra un archivo
***
Function RenameFile(cOldName, cNewName)


If this.Sockets.SendReceive(this.SocketNumber, "RNFR " + cOldName + NEW_LINE)
   If this.Sockets.SocketClosed
      this.ErrorNumber = 4
      Return(.F.)
   else
      this.FTPResponseCode = Val(Left(this.Sockets.Response, 3))
      this.FTPResponseText = SubStr(this.Sockets.Response, 5)

      If (this.FTPResponseCode <> 450) .AND. ;
         (this.FTPResponseCode <> 350)
         this.ErrorNumber = 9
         Return(.F.)
      Endif
   Endif
else
   this.ErrorNumber = 4
   Return(.F.)
Endif


If this.Sockets.SendReceive(this.SocketNumber, "RNTO " + cNewName + NEW_LINE)
   If this.Sockets.SocketClosed
      this.ErrorNumber = 4
      Return(.F.)
   else
      this.FTPResponseCode = Val(Left(this.Sockets.Response, 3))
      this.FTPResponseText = SubStr(this.Sockets.Response, 5)

      If this.FTPResponseCode <> 250
         this.ErrorNumber = 9
         Return(.F.)
      Endif
   Endif
else
   this.ErrorNumber = 4
   Return(.F.)
Endif


Return(.T.)



***
* Function MoveFile(cOldName, cNewName)
*   Mueve un archivo
***
Function MoveFile(cOldName, cNewName)
Return(this.RenameFile(cOldName, cNewName))



***
* Function Dir(cFolder)
*   Return the contents of a folder
***
Function Dir(cFolder)
Local aFiles, nRes, nCont

Dimension aFiles[1]
nRes = this.VFPScriptDir(cFolder, @aFiles)

If nRes > 0
   Dimension this.DirFiles[nRes]
   
   For nCont = 1 to nRes
       this.DirFiles[nCont] = CreateObject("FTPFile")

       this.DirFiles[nCont].FileName = aFiles[nCont, 1]
       this.DirFiles[nCont].Size = aFiles[nCont, 2]
       this.DirFiles[nCont].LastWriteTime = aFiles[nCont, 3]
       this.DirFiles[nCont].RawData = aFiles[nCont, 4]
   Next
Endif

Return(nRes)



***
* Function VFPScriptDir(cFolder, aRes)
*   Return the contents of a folder
***
Function VFPScriptDir(cFolder, aRes)
Local cCommand, cData, nPort, nSocket, cRes, cResult
Local aFiles, nRows, nCont
Local cName, cSize, cDate1, cDate2, cDate3
local cLine, nItem, nStatus, aItems
Local nPos, cLetter, cAntDate
Local nSize, nDay, nMonth, nYear, cHour


* Validar parametros
If Type("cFolder") <> "C"
   cFolder = ""
Endif



* Inicializar propiedades
this.Response = ""


* Ajustar el tipo de transferencia
If this.TransferType <> "A"

   If this.Sockets.SendReceive(this.SocketNumber, "TYPE A" + NEW_LINE)
      If this.Sockets.SocketClosed
         this.ErrorNumber = 4
         Return(-1)
      else
         this.FTPResponseCode = Val(Left(this.Sockets.Response, 3))
         this.FTPResponseText = SubStr(this.Sockets.Response, 5)

         If this.FTPResponseCode <> 200
            this.ErrorNumber = 9
            Return(-1)
         Endif
      Endif
   else
      this.ErrorNumber = 4
      Return(-1)
   Endif

   this.TransferType = "A"

Endif


* Entrar al modo pasivo
If this.Sockets.SendReceive(this.SocketNumber, "PASV" + NEW_LINE)
   If this.Sockets.SocketClosed
      this.ErrorNumber = 4
      Return(-1)
   else
      this.FTPResponseCode = Val(Left(this.Sockets.Response, 3))
      this.FTPResponseText = SubStr(this.Sockets.Response, 5)

      If this.FTPResponseCode <> 227
         this.ErrorNumber = 9
         Return(-1)
      Endif
   Endif
else
   this.ErrorNumber = 4
   Return(-1)
Endif


* Efectuar conexion
nPort = this.GetPort(this.Sockets.Response)
nSocket = this.Sockets.Connect(this.ServerIP, nPort)

If nSocket == 0
   this.ErrorNumber = 10
   Return(-1)
Endif


* Enviar comando
cCommand = "LIST"

If !Empty(cFolder)
   cCommand = cCommand + " " + cFolder
Endif

If this.Sockets.SendReceive(this.SocketNumber, cCommand + NEW_LINE)
   If this.Sockets.SocketClosed
      this.ErrorNumber = 4
      Return(-1)
   else
      this.FTPResponseCode = Val(Left(this.Sockets.Response, 3))
      this.FTPResponseText = SubStr(this.Sockets.Response, 5)

      If this.FTPResponseCode >= 300
         this.ErrorNumber = 9
         Return(-1)
      Endif

      If Occurs(NEW_LINE, this.Sockets.Response) > 1
         cResult = SubStr(this.Sockets.Response, At(NEW_LINE, this.Sockets.Response, Occurs(NEW_LINE, this.Sockets.Response) - 1) + 1)
      Endif

   Endif
else
   this.ErrorNumber = 4
   Return(-1)
Endif


* Esperar respuesta
cData = ""
Do While .T.
   cRes = this.Sockets.Read(nSocket)

   If Len(cRes) <> 0
      cData = cData + cRes
   Endif

   If this.Sockets.SocketClosed

      * Esperar respuesta
      If Empty(cResult)
         If this.Sockets.WaitFor(this.SocketNumber,, 10)
            cResult = this.Sockets.Response
         else
            this.ErrorNumber = 4
            Return(-1)
         Endif
      Endif

      this.FTPResponseCode = Val(Left(cResult, 3))
      this.FTPResponseText = SubStr(cResult, 5)

      If this.FTPResponseCode >= 300
         this.ErrorNumber = 9
         Return(-1)
      else
         Exit
      Endif

   Endif

Enddo

this.Sockets.Close(nSocket)


* Almacenar la respuesta
this.Response = cData


* Procesar la respuesta
Dimension aFiles[1]
nRows = ALines(aFiles, cData)

If nRows <> 0
   Dimension aRes[nRows, 4]
Endif

cAntDate = Set("Date")
Set Date to French

For nCont = 1 to nRows

    * Determinar el formato de la respuesta
    cLine = aFiles[nCont]

    Dimension aItems[1]
    nItem = 0

    For nPos = 1 to Len(cLine)
        cLetter = SubStr(cLine, nPos, 1)

        If cLetter <> " "
           nItem = nItem + 1
           Dimension aItems[nItem]

           If At(" ", SubStr(cLine, nPos)) == 0
              aItems[nItem] = SubStr(cLine, nPos)
              Exit
           else
              aItems[nItem] = SubStr(cLine, nPos, At(" ", SubStr(cLine, nPos)) - 1)
              nPos = nPos + At(" ", SubStr(cLine, nPos)) - 1
            Endif
        Endif
    Next


    If nItem == 4

       * Tipo especial
       cName = aItems[4]
       nSize = Val(aItems[3])

       nDay = Val(SubStr(aItems[1], 4, 2))
       nMonth = Val(Left(aItems[1], 2))
       nYear = Val(Right(aItems[1], 2))

       cHour = aItems[2]
       Do Case
          Case Upper(Right(cHour, 2)) == "PM"
               cHour = PadL(Int(Val(Left(cHour, 2)) + 12), 2, "0") + SubStr(cHour, 3)
               cHour = Left(cHour, Len(cHour) - 2)

          Case Upper(Right(cHour, 2)) == "AM"
               cHour = Left(cHour, Len(cHour) - 2)

       EndCase

    else

       * Unix Type
       cName = ""
       cSize = ""
       cDate1 = ""
       cDate2 = ""
       cDate3 = ""

       nItem = 0
       nStatus = 0

       For nPos = 1 to Len(cLine)
           cLetter = SubStr(cLine, nPos, 1)
           If nStatus == 0
              If cLetter <> " "
                 nItem = nItem + 1
                 nStatus = 1
              Endif
           else
              If cLetter == " "
                 nStatus = 0
              Endif
           Endif

           Do Case
              Case (nItem == 5) .AND. (Len(cSize) == 0)
                   cSize = SubStr(cLine, nPos)

                   If At(" ", cSize) == 0
                      cSize = ""
                   else
                      cSize = Left(cSize, At(" ", cSize) - 1)
                   Endif


              Case (nItem == 6) .AND. (Len(cDate1) == 0)
                   cDate1 = SubStr(cLine, nPos)

                   If At(" ", cDate1) == 0
                      cDate1 = ""
                   else
                      cDate1 = Left(cDate1, At(" ", cDate1) - 1)
                   Endif

              Case (nItem == 7) .AND. (Len(cDate2) == 0)
                   cDate2 = SubStr(cLine, nPos)

                   If At(" ", cDate2) == 0
                      cDate2 = ""
                   else
                      cDate2 = Left(cDate2, At(" ", cDate2) - 1)
                   Endif

              Case (nItem == 8) .AND. (Len(cDate3) == 0)
                   cDate3 = SubStr(cLine, nPos)

                   If At(" ", cDate3) == 0
                      cDate3 = ""
                   else
                      cDate3 = Left(cDate3, At(" ", cDate3) - 1)
                   Endif


              Case nItem == 9
                   cName = SubStr(cLine, nPos)
                   Exit
           EndCase
       Next


       nSize = Val(cSize)

       nDay = Val(cDate2)

       Do Case
          Case Upper(cDate1) == Upper("Jan")
               nMonth = 1

          Case Upper(cDate1) == Upper("Feb")
               nMonth = 2

          Case Upper(cDate1) == Upper("Mar")
               nMonth = 3

          Case Upper(cDate1) == Upper("Apr")
               nMonth = 4

          Case Upper(cDate1) == Upper("May")
               nMonth = 5

          Case Upper(cDate1) == Upper("Jun")
               nMonth = 6

          Case Upper(cDate1) == Upper("Jul")
               nMonth = 7

          Case Upper(cDate1) == Upper("Aug")
               nMonth = 8

          Case Upper(cDate1) == Upper("Sep")
               nMonth = 9

          Case Upper(cDate1) == Upper("Oct")
               nMonth = 10

          Case Upper(cDate1) == Upper("Nov")
               nMonth = 11

          Case Upper(cDate1) == Upper("Dec")
               nMonth = 12

          OtherWise
               nMonth = 12

       EndCase

       If At(":", cDate3) == 0
          nYear = Val(cDate3)
          cHour = "00:00"
       else
          nYear = Year(Date())
          cHour = cDate3
       Endif

    Endif

    aRes[nCont, 1] = cName
    aRes[nCont, 2] = nSize
    aRes[nCont, 3] = CtoT(PadL(Int(nDay), 2, "0") + " " + ;
                          PadL(Int(nMonth), 2, "0") + " " + ;
                          PadL(Int(nYear), 4, "0") + " " + ;
                          cHour)
    aRes[nCont, 4] = cLine

Next

Set Date to (cAntDate)

Return(nRows)



***
* Hidden Function GetPort(cResponse)
*   Determina el puerto al que debemos conectarnos
***
Hidden Function GetPort(cResponse)
Local nPort, cPart1, cPart2

cResponse = SubStr(cResponse, At(",", cResponse, 4) + 1)
cResponse = Left(cResponse, At(")", cResponse) - 1)

cPart1 = PadL(DectoBin(Val(Left(cResponse, At(",", cResponse) - 1))), 8, "0")
cPart2 = PadL(DectoBin(Val(SubStr(cResponse, At(",", cResponse) + 1))), 8, "0")

nPort = BintoDec(cPart1 + cPart2)

Return(nPort)



***
* Function Download
*   Downloads a file
***
Function Download(cSource, cDestination, nStart)
Local lRes

If this.StartDownload(cSource, cDestination, nStart)
   lRes = .T.

   Do While !this.EOT
      If !this.DownloadNextPart()
         lRes = .F.
      Endif
   Enddo
else
   lRes = .F.
Endif

this.EndDownload()

Return(lRes)



***
* Function StartDownload
*   Downloads a file
***
Function StartDownload(cSource, cDestination, nStart)
Local aFiles, nRes, nPort


* Validar parametros
If Type("nStart") <> "N"
   nStart = 0
Endif


* Leer el tamaño del archivo a descargar
Dimension aFiles[1]
nRes = this.VFPScriptDir(cSource, @aFiles)
If nRes == 1
   this.TransferSize = aFiles[1, 2]
else
   this.TransferSize = 0
Endif


* Inicializar propiedades
this.EOT = .F.
this.TransferredBytes = 0
this.TransferSocket = 0
this.TransferResult = ""


* Abrir archivo de destino
If nStart == 0
   this.FileHandle = FCreate(cDestination)
else
   this.FileHandle = FOpen(cDestination, 2)
Endif

If this.FileHandle == -1
   this.ErrorNumber = 101
   Return(.F.)
Endif

If nStart <> 0
   FSeek(this.FileHandle, 0, 2) 
Endif


* Ajustar el tipo de transferencia
If this.TransferType <> "I"

   If this.Sockets.SendReceive(this.SocketNumber, "TYPE I" + NEW_LINE)
      If this.Sockets.SocketClosed
         this.ErrorNumber = 4
         Return(.F.)
      else
         this.FTPResponseCode = Val(Left(this.Sockets.Response, 3))
         this.FTPResponseText = SubStr(this.Sockets.Response, 5)

         If this.FTPResponseCode <> 200
            this.ErrorNumber = 9
            Return(.F.)
         Endif
      Endif
   else
      this.ErrorNumber = 4
      Return(.F.)
   Endif

   this.TransferType = "I"

Endif



* Iniciar transferencia parcial
If nStart <> 0
   If this.Sockets.SendReceive(this.SocketNumber, "REST " + LTrim(Str(nStart, 10, 0)) + NEW_LINE)
      If this.Sockets.SocketClosed
         this.ErrorNumber = 4
         Return(.F.)
      else
         this.FTPResponseCode = Val(Left(this.Sockets.Response, 3))
         this.FTPResponseText = SubStr(this.Sockets.Response, 5)

         If this.FTPResponseCode <> 350
            this.ErrorNumber = 102
            Return(.F.)
         Endif
      Endif
   else
      this.ErrorNumber = 4
      Return(.F.)
   Endif
Endif



* Entrar al modo pasivo
If this.Sockets.SendReceive(this.SocketNumber, "PASV" + NEW_LINE)
   If this.Sockets.SocketClosed
      this.ErrorNumber = 4
      Return(.F.)
   else
      this.FTPResponseCode = Val(Left(this.Sockets.Response, 3))
      this.FTPResponseText = SubStr(this.Sockets.Response, 5)

      If this.FTPResponseCode <> 227
         this.ErrorNumber = 9
         Return(.F.)
      Endif
   Endif
else
   this.ErrorNumber = 4
   Return(.F.)
Endif



* Efectuar conexion
nPort = this.GetPort(this.Sockets.Response)
this.TransferSocket = this.Sockets.Connect(this.ServerIP, nPort)

If this.TransferSocket == 0
   this.ErrorNumber = 10
   Return(.F.)
Endif


* Pedir el archivo
If this.Sockets.SendReceive(this.SocketNumber, "RETR " + cSource + NEW_LINE)
   If this.Sockets.SocketClosed
      this.ErrorNumber = 4
      Return(.F.)
   else
      this.FTPResponseCode = Val(Left(this.Sockets.Response, 3))
      this.FTPResponseText = SubStr(this.Sockets.Response, 5)

      If this.FTPResponseCode >= 300
         this.ErrorNumber = 9
         Return(.F.)
      Endif

      If Occurs(NEW_LINE, this.Sockets.Response) > 1
         this.TransferResult = SubStr(this.Sockets.Response, At(NEW_LINE, this.Sockets.Response, Occurs(NEW_LINE, this.Sockets.Response) - 1) + 1)
      Endif
   Endif
else
   this.ErrorNumber = 4
   Return(.F.)
Endif


Return(.T.)



***
* Function DownloadNextPart
*   Descarga la proxima parte
***
Function DownloadNextPart()
Local cRes, cResult

cRes = this.Sockets.Read(this.TransferSocket)
   
If Len(cRes) <> 0
   If FWrite(this.FileHandle, cRes, Len(cRes)) <> Len(cRes)
      this.ErrorNumber = 103
      Return(.F.)
   Endif

   this.TransferredBytes = this.TransferredBytes + Len(cRes)
Endif


If (this.Sockets.SocketClosed) .OR. ;
   ((this.TransferSize <> 0) .AND. (this.TransferredBytes >= this.TransferSize))

   * Esperar respuesta
   If Empty(this.TransferResult)
      cResult = ""
   else
      cResult = this.TransferResult
   Endif

   this.Sockets.WaitFor(this.SocketNumber,, 3)

   If !Empty(this.Sockets.Response)
      cResult = cResult + this.Sockets.Response
   Endif

   If (this.TransferSize <> 0) .AND. (this.TransferredBytes >= this.TransferSize)
      * No verificar la resputa
   else
      If Empty(cResult)
         this.ErrorNumber = 4
         Return(.F.)
      Endif
   Endif

   this.TransferResult = ""
   this.FTPResponseCode = Val(Left(cResult, 3))
   this.FTPResponseText = SubStr(cResult, 5)


   If (this.TransferSize <> 0) .AND. (this.TransferredBytes >= this.TransferSize)
      this.EOT = .T.
   else
      If this.FTPResponseCode >= 300
         this.ErrorNumber = 9
         Return(.F.)
      else
         this.EOT = .T.
      Endif
   Endif

Endif

Return(.T.)



***
* Function EndDownload
*   Finaliza un download
***
Function EndDownload()

If this.TransferSocket <> 0
   this.Sockets.Close(this.TransferSocket)
   this.TransferSocket = 0
Endif

If this.FileHandle <> -1
   FClose(this.FileHandle)
   this.FileHandle = -1
Endif


EndFunc



***
* Function AbortTransfer
*   Termina una transferencia
***
Function AbortTransfer()
Local cResult

If this.TransferSocket <> 0
   this.Sockets.Close(this.TransferSocket)
   this.TransferSocket = 0


   * Esperar respuesta
   If Empty(this.TransferResult)
      If this.Sockets.WaitFor(this.SocketNumber,, 10)
         cResult = this.Sockets.Response
      else
         this.ErrorNumber = 4
         Return(.F.)
      Endif
   else
      cResult = this.TransferResult
   Endif

   this.TransferResult = ""

   this.FTPResponseCode = Val(Left(cResult, 3))
   this.FTPResponseText = SubStr(cResult, 5)

   If this.FTPResponseCode < 300
      this.ErrorNumber = 9
      Return(.F.)
   Endif

Endif

Return(.T.)



***
* Function Upload
*   Uploads a file
***
Function Upload(cSource, cDestination, lAppend)
Local lRes

If this.StartUpload(cSource, cDestination, lAppend)
   lRes = .T.

   Do While !this.EOT
      If !this.UploadNextPart()
         lRes = .F.
      Endif
   Enddo
else
   lRes = .F.
Endif

this.EndUpload()

Return(lRes)



***
* Function StartUpload
*   Uploads a file
***
Function StartUpload(cSource, cDestination, lAppend)
Local cCommand, nPort


* Inicializar propiedades
this.EOT = .F.
this.TransferredBytes = 0
this.TransferSocket = 0
this.TransferResult = ""


* Abrir archivo de destino
this.FileHandle = FOpen(cSource)

If this.FileHandle == -1
   this.ErrorNumber = 104
   Return(.F.)
Endif


* Ajustar el tipo de transferencia
If this.TransferType <> "I"

   If this.Sockets.SendReceive(this.SocketNumber, "TYPE I" + NEW_LINE)
      If this.Sockets.SocketClosed
         this.ErrorNumber = 4
         this.FTPResponseCode = -1
         this.FTPResponseText = "Este es un código de error interno de iFox.DirectFTP"
         Return(.F.)
      else
         this.FTPResponseCode = Val(Left(this.Sockets.Response, 3))
         this.FTPResponseText = SubStr(this.Sockets.Response, 5)

         If this.FTPResponseCode <> 200
            this.ErrorNumber = 9
            Return(.F.)
         Endif
      Endif
   else
      this.ErrorNumber = 4
      this.FTPResponseCode = -2
      this.FTPResponseText = "Este es un código de error interno de iFox.DirectFTP"
      Return(.F.)
   Endif

   this.TransferType = "I"

Endif


* Entrar al modo pasivo
If this.Sockets.SendReceive(this.SocketNumber, "PASV" + NEW_LINE)
   If this.Sockets.SocketClosed
      this.ErrorNumber = 4
      this.FTPResponseCode = -3
      this.FTPResponseText = "Este es un código de error interno de iFox.DirectFTP"
      Return(.F.)
   else
      this.FTPResponseCode = Val(Left(this.Sockets.Response, 3))
      this.FTPResponseText = SubStr(this.Sockets.Response, 5)

      If this.FTPResponseCode <> 227
         this.ErrorNumber = 9
         Return(.F.)
      Endif
   Endif
else
   this.ErrorNumber = 4
   this.FTPResponseCode = -4
   this.FTPResponseText = "Este es un código de error interno de iFox.DirectFTP"
   Return(.F.)
Endif


* Efectuar conexion
nPort = this.GetPort(this.Sockets.Response)
this.TransferSocket = this.Sockets.Connect(this.ServerIP, nPort)

If this.TransferSocket == 0
   this.ErrorNumber = 10
   Return(.F.)
Endif


* Pedir el archivo
If lAppend
   cCommand = "APPE"
else
   cCommand = "STOR"
Endif

If this.Sockets.SendReceive(this.SocketNumber, cCommand + " " + cDestination + NEW_LINE)
   If this.Sockets.SocketClosed
      this.ErrorNumber = 4
      this.FTPResponseCode = -5
      this.FTPResponseText = "Este es un código de error interno de iFox.DirectFTP"
      Return(.F.)
   else
      this.FTPResponseCode = Val(Left(this.Sockets.Response, 3))
      this.FTPResponseText = SubStr(this.Sockets.Response, 5)

      If this.FTPResponseCode >= 300
         this.ErrorNumber = 105
         Return(.F.)
      Endif
   Endif

   If Occurs(NEW_LINE, this.Sockets.Response) > 1
      this.TransferResult = SubStr(this.Sockets.Response, At(NEW_LINE, this.Sockets.Response, Occurs(NEW_LINE, this.Sockets.Response) - 1) + 1)
   Endif
else
   this.ErrorNumber = 4
   this.FTPResponseCode = -6
   this.FTPResponseText = "Este es un código de error interno de iFox.DirectFTP"
   Return(.F.)
Endif


Return(.T.)



***
* Function UploadNextPart
*   Sube la proxima parte
***
Function UploadNextPart()
Local cData, cResult

If (FEOF(this.FileHandle)) .AND. (!this.EOT)

   * Cerrar transferencia
   If this.TransferSocket <> 0
      this.Sockets.Close(this.TransferSocket)
      this.TransferSocket = 0
   Endif


   * Esperar respuesta
   If Empty(this.TransferResult)
      If this.Sockets.WaitFor(this.SocketNumber,, 10)
         cResult = this.Sockets.Response
      else
         this.ErrorNumber = 4
         Return(.F.)
      Endif
   else
      cResult = this.TransferResult
   Endif

   this.TransferResult = ""

   this.FTPResponseCode = Val(Left(cResult, 3))
   this.FTPResponseText = SubStr(cResult, 5)

   If this.FTPResponseCode >= 300
      this.ErrorNumber = 9
      Return(.F.)
   Endif

   this.EOT = .T.

else
   cData = FRead(this.FileHandle, 4096)
   this.Sockets.Send(this.TransferSocket, cData)
   
   If this.Sockets.SocketClosed
      this.ErrorNumber = 4
      Return(.F.)
   Endif

   this.TransferredBytes = this.TransferredBytes + Len(cData)
Endif

Return(.T.)



***
* Function EndUpload
*   Finaliza un upload
***
Function EndUpload()

If this.FileHandle <> -1
   FClose(this.FileHandle)
   this.FileHandle = -1
Endif


EndFunc



EndDefine



**************************   FTPFILE CLASS   *******************************

****************************************************************************
Define Class FTPFile as Custom

             FileName = ""
             Size = 0
             LastWriteTime = CtoT("")
             RawData = ""

EndDefine
