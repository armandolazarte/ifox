
If .F.
   ?  "Creado por Pablo Pioli"
Endif


#include IFOX.H


#define HTTP_TRANSFER   1
#define FTP_TRANSFER    2


************************   RESUME&GO CLASS   *******************************

****************************************************************************
Define Class ResumeandGo as iFox of iFox.PRG OlePublic

       TransferType = TYPE_FILE
       FileName = ""

       FileSize = 0
       DownloadSize = 0
       DownloadComplete = .F.
       DownloadedBytes = 0
       PreviousDownload = 0

       ErrorNumber = 0
       HTTPErrorNumber = 0

       Hidden Sockets
       Hidden SocketNumber

       Hidden Headers
       Hidden DownloadData

       Hidden Protocol
       Hidden FileHandle


***
* Function eInit
***
Protected Function eInit()


* Inicializar propiedades
this.Protocol = HTTP_TRANSFER
this.SocketNumber = 0
this.FileHandle = -1


EndFunc



***
* Function StartHTTPTransfer
*   Establece un comunicación con el servidor HTTP
***
Function StartHTTPTransfer(cServer, cURL, nStart, nPort, nTimeOut, lDynamic)
Local cIP, cHeaders, cData, tStart, cRes, aHeaders, nPos


* Inicializar propiedades
this.Protocol = HTTP_TRANSFER

this.ErrorNumber = 0
this.HTTPErrorNumber = 0
this.SocketNumber = 0

this.Headers = ""
this.DownloadData = ""

this.FileSize = 0
this.DownloadSize = 0
this.DownloadedBytes = 0
this.PreviousDownload = 0



* Verificar parametros
If Type("cServer") <> "C"
   Return(.F.)
Endif

If Type("cURL") <> "C"
   Return(.F.)
Endif

If Type("nStart") <> "N"
   nStart = 0
Endif

If Type("nPort") <> "N"
   nPort = 80
Endif

If Type("nTimeOut") <> "N"
   nTimeOut = 15
Endif

If (this.TransferType == TYPE_FILE) .AND. (Empty(this.FileName))
   this.ErrorNumber = 8
   Return(.F.)
Endif



* Inicializar socket
this.Sockets = CreateObject("iFox.Sockets")
If !this.Sockets.StartOK
   this.ErrorNumber = 1
   Return(.F.)
Endif


* Abrir archivo
If this.TransferType == TYPE_FILE

   If File(this.FileName)
      this.FileHandle = FOpen(this.FileName, 2)

      If this.FileHandle == -1
         this.ErrorNumber = 9
         Return(.F.)
      Endif
      
      If nStart == 0
         nStart = FSeek(this.FileHandle, 0, 2)
         this.PreviousDownload = nStart
      else
         FSeek(this.FileHandle, 0, 2)
      Endif
   else
      this.FileHandle = FCreate(this.FileName)

      If this.FileHandle == -1
         this.ErrorNumber = 9
         Return(.F.)
      Endif
   Endif

Endif



* Construir encabezado
cIP = this.Sockets.GetIPFromName(cServer)
If Empty(cIP)
   this.ErrorNumber = 2
   Return(.F.)
Endif


cHeaders = "GET " + cURL + " HTTP/1.1" + Chr(13) + Chr(10)
cHeaders = cHeaders + "Host: " + cServer + Chr(13) + Chr(10)
cHeaders = cHeaders + "Accept: */*" + Chr(13) + Chr(10)
cHeaders = cHeaders + "User-Agent: Mozilla/4.0 (compatible; iFox; Windows)" + Chr(13) + Chr(10)
cHeaders = cHeaders + "Range: bytes=" + LTrim(Str(nStart, 10, 0)) + "-" + Chr(13) + Chr(10)
cHeaders = cHeaders + "Pragma: no-cache" + Chr(13) + Chr(10)
cHeaders = cHeaders + "Cache-Control: no-cache" + Chr(13) + Chr(10)
cHeaders = cHeaders + "Connection: close" + Chr(13) + Chr(10)
cHeaders = cHeaders + Chr(13) + Chr(10)


* Establecer conexión
this.SocketNumber = this.Sockets.Connect(cIP, nPort)

If this.SocketNumber == 0
   this.ErrorNumber = 3
   Return(.F.)
Endif


* Enviar encabezado
this.Sockets.Send(this.SocketNumber, cHeaders)


* Esperar respuesta
cData = ""

tStart = DateTime()

Do While .T.

   cRes = this.Sockets.Read(this.SocketNumber)

   If Len(cRes) <> 0
      cData = cData + cRes
   Endif
   
   If At(NEW_LINE + NEW_LINE, cData) <> 0
      this.Headers = Left(cData, At(NEW_LINE + NEW_LINE, cData) - 1)
      this.DownLoadData = SubStr(cData, At(NEW_LINE + NEW_LINE, cData) + Len(NEW_LINE + NEW_LINE))

      Dimension aHeaders[1]
      If ALines(aHeaders, this.Headers) < 1
         this.ErrorNumber = 4
         Return(.F.)
      Endif


      If Left(aHeaders[1], 8) <> "HTTP/1.1"
         this.ErrorNumber = 5
         Return(.F.)
      Endif


      If lDynamic

         If SubStr(aHeaders[1], 10, 3) <> "200"
            this.ErrorNumber = 10
            this.HTTPErrorNumber = Val(SubStr(aHeaders[1], 10, 3))
            Return(.F.)
         Endif

         For nPos = 1 to ALen(aHeaders)
             If Upper(Left(aHeaders[nPos], 15)) == Upper("Content-Length:")
                this.DownloadSize = Val(SubStr(aHeaders[nPos], 17))
                this.FileSize = this.DownloadSize
                Exit
             Endif
         Next

      else

         If SubStr(aHeaders[1], 10, 3) == "200"
            this.ErrorNumber = 7
            Return(.F.)
         Endif

         If SubStr(aHeaders[1], 10, 3) <> "206"
            this.ErrorNumber = 6
            this.HTTPErrorNumber = Val(SubStr(aHeaders[1], 10, 3))
            Return(.F.)
         Endif

         For nPos = 1 to ALen(aHeaders)
             If Upper(Left(aHeaders[nPos], 14)) == Upper("Content-Range:")
                this.FileSize = Val(SubStr(aHeaders[nPos], Rat("/", aHeaders[nPos]) + 1))
                Exit
             Endif
         Next

         If nPos == ALen(aHeaders) + 1
            this.ErrorNumber = 10
            Return(.F.)
         Endif

         For nPos = 1 to ALen(aHeaders)
             If Upper(Left(aHeaders[nPos], 15)) == Upper("Content-Length:")
                this.DownloadSize = Val(SubStr(aHeaders[nPos], 17))
                Exit
             Endif
         Next

      Endif


      If nPos == ALen(aHeaders) + 1
         this.ErrorNumber = 10
         Return(.F.)
      Endif


      * Guardar lo recibido hasta el momento
      If this.TransferType == TYPE_FILE
         If FWrite(this.FileHandle, this.DownLoadData) <> Len(this.DownLoadData)
            this.ErrorNumber = 102
            Return(.F.)
         Endif
      Endif

      this.DownloadedBytes = Len(this.DownLoadData)


      Exit
   Endif

   If DateTime() - tStart > nTimeOut
      Exit
   Endif

EndDo


EndFunc



***
* Hidden Function DownloadComplete_Access
*   Determina si se ha terminado la transferencia
***
Hidden Function DownloadComplete_Access()
Local lRes

If this.DownloadedBytes >= this.DownloadSize
   lRes = .T.
else
   lRes = .F.
Endif

Return(lRes)



***
* Function DownloadNextPart
*   Descarga la proxima parte
***
Function DownloadNextPart()
Local lRes, cRes

cRes = this.Sockets.Read(this.SocketNumber)

If (this.Sockets.SocketClosed) .AND. (Len(cRes) == 0)
   lRes = .F.
   this.ErrorNumber = 101
else
   lRes = .T.
   If Len(cRes) <> 0

      If this.TransferType == TYPE_FILE
         If FWrite(this.FileHandle, cRes) <> Len(cRes)
            lRes = .F.
            this.ErrorNumber = 102
         Endif
      else
         this.DownLoadData = this.DownLoadData + cRes
      Endif

      this.DownloadedBytes = this.DownloadedBytes + Len(cRes)
   Endif
Endif

Return(lRes)



***
* Function GetZeroCount
*   Determina si la cadena en memoria comienza en Chr(0)
***
Function GetZeroCount()
Local nPos

nPos = 1
Do While SubStr(this.DownloadData, nPos, 1) == Chr(0)
   nPos = nPos + 1
Enddo

Return(nPos - 1)



***
* Function GetPartialDownload
*   Limpia el buffer de resultado
***
Function GetPartialDownload()
Local nZero

nZero = this.GetZeroCount()

If nZero <> 0
   Return(SubStr(this.DownloadData, nZero + 1))
else
   Return(this.DownloadData)
Endif

EndFunc



***
* Function ClearDownLoadData
*   Limpia el buffer de resultado
***
Function ClearDownLoadData()
this.DownloadData = ""
EndFunc



***
* Function EndTransfer
*   Cierra la comunicación con el servidor HTTP
***
Function EndTransfer()

If this.Protocol == HTTP_TRANSFER
   If this.SocketNumber <> 0
      this.Sockets.Close(this.SocketNumber)
   Endif
else
Endif

If this.TransferType == TYPE_FILE
   If this.FileHandle <> -1
      FClose(this.FileHandle)
   Endif
Endif

EndFunc


EndDefine
