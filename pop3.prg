
If .F.
   ?  "Creado por Pablo Pioli"
Endif


#include IFOX.H


***************************   POP3 CLASS   **********************************

****************************************************************************
Define Class POP3 as iFox of iFox.PRG OlePublic

       Hidden SocketsHandler
       Hidden ConnectionSocket

       UserName = ""
       Password = ""

       MessageCount = 0
       TotalSize = 0

       DownloadHeaders = .F.
       DownloadLines = 10

       ErrorNumber = 0
       POP3Response = ""

       AsyncMessage = ""
       EOD = .F.
       Transferred = 0

       RawHeaders = ""
       RawBody = ""
       RawMessage = ""

       AttachmentsCount = 0
       Dimension Attachments[1]

       ResourcesCount = 0
       Dimension Resources[1]

       ListCount = 0
       Dimension ListData[1]


       * Headers del mensaje
       SenderName = ""
       SenderEMail = ""

       RecipientsCount = 0
       CCCount = 0

       Dimension Recipients[1]
       Dimension CC[1]

       Subject = ""
       Priority = SMTP_PRIORITY_NORMAL

       SentDate = CtoD("")
       SentTime = ""
       SentTimeZone = 0
       AdjustedDate = CtoD("")
       AdjustedTime = ""

       Body = ""
       HTMLBody = ""


       * Auxiliares a la decodificacion del mensaje
       Hidden HeadersCount
       Hidden HeadersData[1]


***
* Function eInit
***
Protected Function eInit()


* Inicializar propiedades
this.SocketsHandler = CreateObject("Sockets")
this.ConnectionSocket = 0


EndFunc



***
* Function Connect
***
Function Connect(cServer, nPort)
Local lRes, cResponse


* Validar puerto
If (Type("nPort") <> "N") .OR. (nPort == 0)
   nPort = 110
Endif


* Inicializar propiedades
this.MessageCount = 0
this.TotalSize = 0
this.POP3Response = ""


* Establecer conexion
this.ConnectionSocket = this.SocketsHandler.Connect(cServer, nPort)

If this.ConnectionSocket <> 0
   lRes = .T.
else
   this.ErrorNumber = 1
   lRes = .F.
Endif


* Esperar saludo
If lRes
   If this.SocketsHandler.WaitFor(this.ConnectionSocket, "", 10)
      If Left(this.SocketsHandler.Response, 1) <> "+"
         this.ErrorNumber = 3
         lRes = .F.
      Endif
   else
      this.ErrorNumber = 1
      lRes = .F.
   Endif
Endif


* Enviar nombre de usuario
If lRes
   If this.SocketsHandler.SendReceive(this.ConnectionSocket, "USER " + this.UserName + NEW_LINE, 10)
      If Left(this.SocketsHandler.Response, 1) <> "+"
         this.ErrorNumber = 12
         lRes = .F.
      Endif
   else
      this.ErrorNumber = 2
      lRes = .F.
   Endif
Endif


* Enviar clave de acceso
If lRes
   If this.SocketsHandler.SendReceive(this.ConnectionSocket, "PASS " + this.Password + NEW_LINE, 10)
      If Left(this.SocketsHandler.Response, 1) <> "+"
         this.ErrorNumber = 12
         lRes = .F.
      Endif
   else
      this.ErrorNumber = 2
      lRes = .F.
   Endif
Endif


* Pedir estado
If lRes
   If this.SocketsHandler.SendReceive(this.ConnectionSocket, "STAT" + NEW_LINE, 10)
      If Left(this.SocketsHandler.Response, 1) == "+"

         cResponse = this.SocketsHandler.Response

         cResponse = SubStr(cResponse, At(" ", cResponse) + 1)
         this.MessageCount = Val(Left(cResponse, At(" ", cResponse) - 1))

         cResponse = SubStr(cResponse, At(" ", cResponse) + 1)
         If At(" ", cResponse) == 0
            this.TotalSize = Val(cResponse)
         else
            this.TotalSize = Val(Left(cResponse, At(" ", cResponse) - 1))
         Endif

      else
         this.ErrorNumber = 3
         lRes = .F.
      Endif
   else
      this.ErrorNumber = 2
      lRes = .F.
   Endif
Endif


* Desconectar
If !lRes
   this.Disconnect()
Endif


Return(lRes)



***
* Function DisConnect
***
Function DisConnect()
Local lRes

lRes = .T.

If this.ConnectionSocket <> 0

   If this.SocketsHandler.SendReceive(this.ConnectionSocket, "QUIT" + NEW_LINE, 10)
      If Left(this.SocketsHandler.Response, 1) <> "+"
         this.ErrorNumber = 3
         lRes = .F.
      Endif
   else
      this.ErrorNumber = 2
      lRes = .F.
   Endif

   this.SocketsHandler.Close(this.ConnectionSocket)
else
   this.ErrorNumber = 1
   lRes = .F.
Endif

Return(lRes)



***
* Function Delete
***
Function Delete(nMessage)
Local lRes

lRes = .T.
this.POP3Response = ""

If this.ConnectionSocket <> 0
   If this.SocketsHandler.SendReceive(this.ConnectionSocket, "DELE " + LTrim(Str(Int(nMessage))) + NEW_LINE, 10)
      this.POP3Response = this.SocketsHandler.Response

      If Left(this.SocketsHandler.Response, 1) <> "+"
         this.ErrorNumber = 3
         lRes = .F.
      Endif
   else
      this.ErrorNumber = 2
      lRes = .F.
   Endif
else
   this.ErrorNumber = 1
   lRes = .F.
Endif

Return(lRes)



***
* Function Recall
***
Function Recall()
Local lRes

lRes = .T.
this.POP3Response = ""

If this.ConnectionSocket <> 0
   If this.SocketsHandler.SendReceive(this.ConnectionSocket, "RSET" + NEW_LINE, 10)
      this.POP3Response = this.SocketsHandler.Response

      If Left(this.SocketsHandler.Response, 1) <> "+"
         this.ErrorNumber = 3
         lRes = .F.
      Endif
   else
      this.ErrorNumber = 2
      lRes = .F.
   Endif
else
   this.ErrorNumber = 1
   lRes = .F.
Endif

Return(lRes)



***
* Function Test
***
Function Test()
Local lRes

lRes = .T.
this.POP3Response = ""

If this.ConnectionSocket <> 0
   If this.SocketsHandler.SendReceive(this.ConnectionSocket, "NOOP" + NEW_LINE, 10)
      this.POP3Response = this.SocketsHandler.Response

      If Left(this.SocketsHandler.Response, 1) <> "+"
         this.ErrorNumber = 3
         lRes = .F.
      Endif
   else
      this.ErrorNumber = 2
      lRes = .F.
   Endif
else
   this.ErrorNumber = 1
   lRes = .F.
Endif

Return(lRes)



***
* Function List
***
Function List()
Local lRes, cCommand, cResponse, cData
Local nLines, aMessages, nCont, oMessage

lRes = .T.
cResponse = ""

this.POP3Response = ""

If this.ConnectionSocket <> 0
   If this.SocketsHandler.SendReceive(this.ConnectionSocket, "LIST" + NEW_LINE, 10)
      this.POP3Response = this.SocketsHandler.Response

      If Left(this.SocketsHandler.Response, 1) == "+"
         cResponse = this.SocketsHandler.Response

         Do While .T.
            cData = this.SocketsHandler.Read(this.ConnectionSocket)

            If Len(cData) <> 0
               cResponse = cResponse + cData
            Endif

            If At(NEW_LINE + "." + NEW_LINE, cResponse) <> 0
               cResponse = Left(cResponse, At(NEW_LINE + "." + NEW_LINE, cResponse) - 1)
               cResponse = SubStr(cResponse, At(NEW_LINE, cResponse) + 2)

               lRes = .T.
               Exit
            Endif


            If this.SocketsHandler.SocketClosed
               this.ErrorNumber = 2
               lRes = .F.
               Exit
            Endif

         Enddo
      else
         this.ErrorNumber = 3
         lRes = .F.
      Endif

   else
      this.ErrorNumber = 2
      lRes = .F.
   Endif
else
   this.ErrorNumber = 1
   lRes = .F.
Endif

If lRes
   Dimension aMessages[1]
   nLines = ALines(aMessages, cResponse, .T., NEW_LINE)

   this.ListCount = 0
   For nCont = 1 to nLines
       If Left(aMessages[nCont], 1) <> "+"

          this.ListCount = this.ListCount + 1
          Dimension this.ListData[this.ListCount]

          oMessage = CreateObject("MailMessage")
          oMessage.Number = Val(Left(aMessages[nCont], At(" ", aMessages[nCont]) - 1))
          oMessage.Size = Val(SubStr(aMessages[nCont], At(" ", aMessages[nCont]) + 1))
          this.ListData[this.ListCount] = oMessage

       Endif
   Next
Endif

Return(lRes)



***
* Function Get
***
Function Get(nMessage)
Local lRes, cCommand, cMessage, cData


If this.DownloadHeaders
   cCommand = "TOP " + LTrim(Str(Int(nMessage))) + " " + ;
                        LTrim(Str(Int(this.DownloadLines)))
else
   cCommand = "RETR " + LTrim(Str(Int(nMessage)))
Endif


lRes = .T.
cMessage = ""

this.POP3Response = ""

If this.ConnectionSocket <> 0
   If this.SocketsHandler.SendReceive(this.ConnectionSocket, cCommand + NEW_LINE, 10)
      this.POP3Response = this.SocketsHandler.Response

      If Left(this.SocketsHandler.Response, 1) == "+"
         cMessage = this.SocketsHandler.Response

         Do While .T.
            cData = this.SocketsHandler.Read(this.ConnectionSocket)

            If Len(cData) <> 0
               cMessage = cMessage + cData
            Endif


            If At(NEW_LINE + "." + NEW_LINE, cMessage) <> 0
               cMessage = Left(cMessage, At(NEW_LINE + "." + NEW_LINE, cMessage) - 1)
               cMessage = SubStr(cMessage, At(NEW_LINE, cMessage) + 2)

               lRes = .T.
               Exit
            Endif


            If this.SocketsHandler.SocketClosed
               this.ErrorNumber = 2
               lRes = .F.
               Exit
            Endif

         Enddo
      else
         this.ErrorNumber = 3
         lRes = .F.
      Endif

   else
      this.ErrorNumber = 2
      lRes = .F.
   Endif
else
   this.ErrorNumber = 1
   lRes = .F.
Endif

If lRes
   If (this.DownloadHeaders) .AND. (this.DownloadLines == 0)
      cMessage = cMessage + NEW_LINE
   Endif

   lRes = this.ParseMessage(cMessage)
Endif

Return(lRes)



***
* Function StartAsyncGet
***
Function StartAsyncGet(nMessage)
Local lRes, cCommand


If this.DownloadHeaders
   cCommand = "TOP " + LTrim(Str(Int(nMessage))) + " " + ;
                        LTrim(Str(Int(this.DownloadLines)))
else
   cCommand = "RETR " + LTrim(Str(Int(nMessage)))
Endif


lRes = .T.

this.AsyncMessage = ""
this.EOD = .F.
this.Transferred = 0

this.POP3Response = ""

If this.ConnectionSocket <> 0
   If this.SocketsHandler.SendReceive(this.ConnectionSocket, cCommand + NEW_LINE, 10)
      this.POP3Response = this.SocketsHandler.Response

      If Left(this.SocketsHandler.Response, 1) == "+"
         this.AsyncMessage = this.SocketsHandler.Response
      else
         this.ErrorNumber = 3
         lRes = .F.
      Endif

   else
      this.ErrorNumber = 2
      lRes = .F.
   Endif
else
   this.ErrorNumber = 1
   lRes = .F.
Endif


Return(lRes)



***
* Function ContinueAsyncGet
***
Function ContinueAsyncGet()
Local lRes, cData

lRes = .T.
cData = this.SocketsHandler.Read(this.ConnectionSocket)

If Len(cData) <> 0
   this.AsyncMessage = this.AsyncMessage + cData
   this.Transferred = Len(this.AsyncMessage)
Endif

If At(NEW_LINE + "." + NEW_LINE, this.AsyncMessage) <> 0
   this.AsyncMessage = Left(this.AsyncMessage, At(NEW_LINE + "." + NEW_LINE, this.AsyncMessage) - 1)
   this.AsyncMessage = SubStr(this.AsyncMessage, At(NEW_LINE, this.AsyncMessage) + 2)
   this.EOD = .T.
   Return(.T.)
Endif

If this.SocketsHandler.SocketClosed
   this.ErrorNumber = 2
   lRes = .F.
Endif

Return(lRes)



***
* Function EndAsyncGet
***
Function EndAsyncGet()
Local lRes

If (this.DownloadHeaders) .AND. (this.DownloadLines == 0)
   this.AsyncMessage = this.AsyncMessage + NEW_LINE
Endif

lRes = this.ParseMessage(this.AsyncMessage)

this.AsyncMessage = ""

Return(lRes)



***
* Function ParseMessage
***
Function ParseMessage(cMssg)
Local nSeparator, nCont, cContent, cBoundary
Local oUtils, cBody, cEncoding


* Determinar donde esta el separador
nSeparator = At(NEW_LINE + NEW_LINE, cMssg)

If nSeparator == 0 then
   this.RawHeaders = cMssg
   this.RawBody = ""

   this.ErrorNumber = 4
   Return(.F.)
Endif


* Separar los items en "bruto"
this.RawHeaders = Left(cMssg, nSeparator - 1)
this.RawBody = SubStr(cMssg, nSeparator + Len(NEW_LINE + NEW_LINE))


* Separar los items
this.HeadersCount = ALines(this.HeadersData, Left(cMssg, nSeparator), .F., Chr(13))


* Limpiar los encabezados
For nCont = 1 to this.HeadersCount
    this.HeadersData[nCont] = StrTran(this.HeadersData[nCont], Chr(13), "")
    this.HeadersData[nCont] = StrTran(this.HeadersData[nCont], Chr(10), "")
Next


* Extraer los encabezados
this.ProcessHeaders()


* Determinar el divisor usado
cContent = this.GetHeader("Content-Type:")

If Upper(Left(cContent, Len("multipart"))) == Upper("multipart")
   cBoundary = SubStr(cContent, ;
               At(Upper("boundary="), Upper(cContent)) + Len("boundary="))

   If Left(cBoundary, 1) == '"'
      cBoundary = SubStr(cBoundary, 2)
      cBoundary = Left(cBoundary, At('"', cBoundary) - 1)
   Endif
else
   cBoundary = ""
Endif


* Procesar el cuerpo
If Len(cBoundary) == 0

   oUtils = CreateObject("Utils")

   cBody = this.RawBody
   cEncoding = this.GetHeader("Content-Transfer-Encoding:")

   Do Case
      Case Upper(cEncoding) == Upper("quoted-printable")
           cBody = oUtils.Decode(cBody, 1)

      Case Upper(cEncoding) == Upper("base64")
           cBody = StrConv(cBody, 14)

   EndCase


   If Upper(Left(cContent, Len("text/html"))) == Upper("text/html")
      this.Body = ""
      this.HTMLBody = cBody
   else
      this.Body = cBody
      this.HTMLBody = ""
   Endif


   this.AttachmentsCount = 0
   Dimension this.Attachments[1]
   this.Attachments[1, 1] = .F.

   this.ResourcesCount = 0
   Dimension this.Resources[1]
   this.Resources[1] = .F.

else
   this.DecodeMessage(this.RawBody, cBoundary)
Endif


Return(.T.)



***
* Function ProcessHeaders()
***
Hidden Function ProcessHeaders()
Local oUtils, cHeader, cName, cEMail


* Inicializar el objeto auxiliar
oUtils = CreateObject("Utils")


* Leer remitente
cHeader = this.GetHeader("From:")
If RAT("<", cHeader) <> 0
   this.SenderName = oUtils.Decode(AllTrim(Left(cHeader, RAT("<", cHeader) - 1)), 3)
   If (Left(this.SenderName, 1) == '"') .AND. (Right(this.SenderName, 1) == '"')
      this.SenderName = SubStr(this.SenderName, 2, Len(this.SenderName) - 2)
   Endif

   cHeader = AllTrim(SubStr(cHeader, AT("<", cHeader) + 1))
   cHeader = StrTran(cHeader, ">", "")
   this.SenderEMail = oUtils.Decode(cHeader, 3)
else
   this.SenderName = ""
   this.SenderEMail = oUtils.Decode(cHeader, 3)
Endif


* Leer datos auxiliares
this.Subject = oUtils.Decode(this.GetHeader("Subject:"), 3)
this.Priority = Val(oUtils.Decode(this.GetHeader("X-Priority:"), 3))
this.DecodeDate(oUtils.Decode(this.GetHeader("Date:"), 3))


* Leer destinatarios
this.RecipientsCount = 0
Dimension this.Recipients[1]

cHeader = StrTran(oUtils.Decode(this.GetHeader("To:"), 3), Chr(13), "")

Do While !Empty(cHeader)
   If (AT("<", cHeader) == 0) .OR. (AT(">", cHeader) == 0)
      cName = ""
      If At(",", cHeader) <> 0
         cEMail = AllTrim(Left(cHeader, At(",", cHeader) - 1))
      else
         cEMail = AllTrim(cHeader)
      Endif
   else
      cName = AllTrim(Left(cHeader, AT("<", cHeader) - 1))
      If (Left(cName, 1) == '"') .AND. (Right(cName, 1) == '"')
         cName = SubStr(cName, 2, Len(cName) - 2)
      Endif

      cEMail = AllTrim(SubStr(cHeader, AT("<", cHeader) + 1))
      cEMail = Left(cEMail, At(">", cEMail) - 1)
      cEMail = cEMail
   Endif

   this.RecipientsCount = this.RecipientsCount + 1
   Dimension this.Recipients[this.RecipientsCount]
   this.Recipients[this.RecipientsCount] = CreateObject("MailReceiver")
   this.Recipients[this.RecipientsCount].ReceiverName = cName
   this.Recipients[this.RecipientsCount].ReceiverEMail = cEMail

   If At(">", cHeader) <> 0
      cHeader = AllTrim(SubStr(cHeader, At(">", cHeader) + 1))
   else
      If At(",", cHeader) <> 0
         cHeader = SubStr(cHeader, At(",", cHeader))
      else
         cHeader = ""
      Endif
   Endif

   If Left(cHeader, 1) == ","
      cHeader = SubStr(cHeader, 2)
   else
      Exit
   Endif
Enddo


* Leer copias
this.CCCount = 0
Dimension this.CC[1]

cHeader = StrTran(oUtils.Decode(this.GetHeader("CC:"), 3), Chr(13), "")

Do While !Empty(cHeader)
   If (AT("<", cHeader) == 0) .OR. (AT(">", cHeader) == 0)
      cName = ""
      If At(",", cHeader) <> 0
         cEMail = AllTrim(Left(cHeader, At(",", cHeader) - 1))
      else
         cEMail = AllTrim(cHeader)
      Endif
   else
      cName = AllTrim(Left(cHeader, AT("<", cHeader) - 1))
      If (Left(cName, 1) == '"') .AND. (Right(cName, 1) == '"')
         cName = SubStr(cName, 2, Len(cName) - 2)
      Endif

      cEMail = AllTrim(SubStr(cHeader, AT("<", cHeader) + 1))
      cEMail = Left(cEMail, At(">", cEMail) - 1)
      cEMail = cEMail
   Endif

   this.CCCount = this.CCCount + 1
   Dimension this.CC[this.CCCount]
   this.CC[this.CCCount] = CreateObject("MailReceiver")
   this.CC[this.CCCount].ReceiverName = cName
   this.CC[this.CCCount].ReceiverEMail = cEMail

   If At(">", cHeader) <> 0
      cHeader = AllTrim(SubStr(cHeader, At(">", cHeader) + 1))
   else
      If At(",", cHeader) <> 0
         cHeader = SubStr(cHeader, At(",", cHeader))
      else
         cHeader = ""
      Endif
   Endif

   If Left(cHeader, 1) == ","
      cHeader = SubStr(cHeader, 2)
   else
      Exit
   Endif
Enddo


EndFunc



***
* Function GetHeader
***
Function GetHeader(cItem)
Local nCont, cRes

cRes = ""
For nCont = 1 to this.HeadersCount
    If Upper(Left(this.HeadersData[nCont], Len(cItem))) == Upper(cItem)

       cRes = AllTrim(this.HeadersData[nCont])

       Do While .T.
          nCont = nCont + 1

          If nCont > this.HeadersCount
             Exit
          Endif

          If (Left(this.HeadersData[nCont], 1) <> " ") .AND. ;
             (Left(this.HeadersData[nCont], 1) <> Chr(9))
             Exit
          Endif

          cRes = cRes + Chr(13) + AllTrim(StrTran(this.HeadersData[nCont], Chr(9), ""))
       Enddo

    Endif
Next

If Upper(Left(cRes, Len(cItem))) == Upper(cItem)
   cRes = SubStr(cRes, Len(cItem) + 1)
Endif

cRes = AllTrim(cRes)

If Left(cRes, 1) == Chr(9)
   cRes = SubStr(cRes, 2)
Endif

Return(cRes)



***
* Function DecodeMessage
***
Hidden Function DecodeMessage(cBody, cBoundary)
Local oUtils, cPart, nCont, nSeparator
Local cContent, cEncoding, cDisposition, cFileName
Local cPartBoundary, cID
Local oAttachment, oResource


* Crear objeto auxiliar
oUtils = CreateObject("Utils")


* Preparar el entorno
this.Body = ""
this.HTMLBody = ""

this.AttachmentsCount = 0
Dimension this.Attachments[1]
this.Attachments[1, 1] = .F.

this.ResourcesCount = 0
Dimension this.Resources[1]
this.Resources[1] = .F.

cBoundary = "--" + cBoundary


* Extraer los submensajes
If At(cBoundary, cBody) <> 0


   * Eliminar la primera parte
   cBody = SubStr(cBody, At(cBoundary, cBody) + Len(cBoundary))
   If Left(cBody, Len(NEW_LINE)) == NEW_LINE
      cBody = SubStr(cBody, Len(NEW_LINE) + 1)
   Endif


   * Extraer el resto
   Do While At(cBoundary, cBody) <> 0

      * Extraer la seccion
      cPart = Left(cBody, At(cBoundary, cBody) - 1)


      * Buscar los encabezados
       nSeparator = At(NEW_LINE + NEW_LINE, cPart)
       If nSeparator <> 0 then


          * Extraer los encabezados
          this.HeadersCount = ALines(this.HeadersData, Left(cPart, nSeparator), .F., Chr(13))
          cPart = SubStr(cPart, nSeparator + Len(NEW_LINE + NEW_LINE))


          * Limpiar los encabezados
          For nCont = 1 to this.HeadersCount
              this.HeadersData[nCont] = StrTran(this.HeadersData[nCont], Chr(13), "")
              this.HeadersData[nCont] = StrTran(this.HeadersData[nCont], Chr(10), "")
          Next


          * Ver si esta seccion no contiene subsecciones
          cContent = this.GetHeader("Content-Type:")
          If Upper(Left(cContent, Len("multipart"))) == Upper("multipart")

             cPartBoundary = SubStr(cContent, ;
                             At(Upper("boundary="), Upper(cContent)) + Len("boundary="))

             If Left(cPartBoundary, 1) == '"'
                cPartBoundary = SubStr(cPartBoundary, 2)
                cPartBoundary = Left(cPartBoundary, At('"', cPartBoundary) - 1)
             Endif

             this.DecodeMessage(cPart, cPartBoundary)

          else

             * Decodificar
             cEncoding = this.GetHeader("Content-Transfer-Encoding:")

             Do Case
                Case Upper(cEncoding) == Upper("quoted-printable")
                     cPart = oUtils.Decode(cPart, 1)

                Case Upper(cEncoding) == Upper("base64")
                     cPart = StrConv(cPart, 14)

             EndCase


             * Guardar el mensaje
             cDisposition = this.GetHeader("Content-Disposition:")

             Do Case
                Case Upper(Left(cDisposition, Len("attachment"))) == Upper("attachment")
                     cFileName = SubStr(cDisposition, At(Upper("FileName="), Upper(cDisposition)) + Len("FileName="))
                     If Left(cFileName, 1) == '"'
                        cFileName = SubStr(cFileName, 2)
                        cFileName = Left(cFileName, At('"', cFileName) - 1)
                     Endif

                     this.AttachmentsCount = this.AttachmentsCount + 1
                     Dimension this.Attachments[this.AttachmentsCount]

                     oAttachment = CreateObject("MailAttachment", cPart)
                     oAttachment.FileName = oUtils.Decode(cFileName, 3)
                     oAttachment.FileSize = Len(cPart)
                     this.Attachments[this.AttachmentsCount] = oAttachment


                     cID = this.GetHeader("Content-ID:")
                     If !Empty(cID)

                        If Left(cID, 1) == "<"
                           cID = SubStr(cID, 2)
                           cID = Left(cID, At(">", cID) - 1)
                        Endif

                        this.ResourcesCount = this.ResourcesCount + 1
                        Dimension this.Resources[this.ResourcesCount,1]

                        oResource = CreateObject("MailResource")
                        oResource.FileName = cFileName
                        oResource.Content = cPart
                        oResource.ID = cID
                        this.Resources[this.ResourcesCount] = oResource

                     Endif


                Case Upper(Left(cContent, Len("text/plain"))) == Upper("text/plain")
                     If Len(this.Body) <> 0
                        this.Body = this.Body + NEW_LINE + NEW_LINE + NEW_LINE
                     Endif
                     this.Body = cPart


                     If At(Upper("FileName="), Upper(cDisposition)) <> 0
                        cFileName = SubStr(cDisposition, At(Upper("FileName="), Upper(cDisposition)) + Len("FileName="))
                        If Left(cFileName, 1) == '"'
                           cFileName = SubStr(cFileName, 2)
                           cFileName = Left(cFileName, At('"', cFileName) - 1)
                        Endif

                        this.AttachmentsCount = this.AttachmentsCount + 1
                        Dimension this.Attachments[this.AttachmentsCount]

                        oAttachment = CreateObject("MailAttachment", cPart)
                        oAttachment.FileName = oUtils.Decode(cFileName, 3)
                        oAttachment.FileSize = Len(cPart)
                        this.Attachments[this.AttachmentsCount] = oAttachment
                     Endif


                Case Upper(Left(cContent, Len("text/html"))) == Upper("text/html")
                     If Len(this.HTMLBody) <> 0
                        this.HTMLBody = this.HTMLBody + "<br><hr><br>"
                     Endif
                     this.HTMLBody = this.HTMLBody + cPart


                     If At(Upper("FileName="), Upper(cDisposition)) <> 0
                        cFileName = SubStr(cDisposition, At(Upper("FileName="), Upper(cDisposition)) + Len("FileName="))
                        If Left(cFileName, 1) == '"'
                           cFileName = SubStr(cFileName, 2)
                           cFileName = Left(cFileName, At('"', cFileName) - 1)
                        Endif

                        this.AttachmentsCount = this.AttachmentsCount + 1
                        Dimension this.Attachments[this.AttachmentsCount]

                        oAttachment = CreateObject("MailAttachment", cPart)
                        oAttachment.FileName = oUtils.Decode(cFileName, 3)
                        oAttachment.FileSize = Len(cPart)
                        this.Attachments[this.AttachmentsCount] = oAttachment
                     Endif


                Case Upper(Left(cContent, Len("message/rfc822"))) == Upper("message/rfc822")
                     this.AttachmentsCount = this.AttachmentsCount + 1
                     Dimension this.Attachments[this.AttachmentsCount]

                     oAttachment = CreateObject("MailAttachment", cPart)
                     oAttachment.FileName = "Attached_Message.EML"
                     oAttachment.FileSize = Len(cPart)
                     this.Attachments[this.AttachmentsCount] = oAttachment


                OtherWise
                     cFileName = SubStr(cContent, At(Upper("Name="), Upper(cContent)) + Len("Name="))
                     If Left(cFileName, 1) == '"'
                        cFileName = SubStr(cFileName, 2)
                        cFileName = Left(cFileName, At('"', cFileName) - 1)
                     Endif


                     cID = this.GetHeader("Content-ID:")
                     If Left(cID, 1) == "<"
                        cID = SubStr(cID, 2)
                        cID = Left(cID, At(">", cID) - 1)
                     Endif


                     this.ResourcesCount = this.ResourcesCount + 1
                     Dimension this.Resources[this.ResourcesCount,1]

                     oResource = CreateObject("MailResource")
                     oResource.FileName = oUtils.Decode(cFileName, 3)
                     oResource.Content = cPart
                     oResource.ID = cID
                     this.Resources[this.ResourcesCount] = oResource

             EndCase

          Endif
       Endif


      * Seguir con la siguiente seccion
      cBody = SubStr(cBody, At(cBoundary, cBody) + Len(cBoundary))
      If Left(cBody, Len(NEW_LINE)) == NEW_LINE
         cBody = SubStr(cBody, Len(NEW_LINE) + 1)
      Endif
   Enddo

Endif

EndFunc



***
* Function DecodeDate
***
Hidden Function DecodeDate(cDate)
Local dDate, cTime, nTimeZone, cAux
Local cPart, nDay, nMonth, nYear
Local oUtils, nCurrentTimeZone, nDif

dDate = CtoD("")
cTime = ""
nTimeZone = 0

If At(",", cDate) <> 0
   If IsDigit(Right(SubStr(cDate, 1, At(",", cDate) - 1), 1))
      cAux = SubStr(cDate, 1, At(",", cDate) - 1)
      cDate = SubStr(cDate, At(",", cDate) + 1)
      Do While Len(cAux) > 0
         If IsDigit(Right(cAux, 1))
            cDate = Right(cAux, 1) + cDate
         else
            Exit
         Endif
         cAux = Left(cAux, Len(cAux) - 1)
      Enddo
   else
     cDate = AllTrim(SubStr(cDate, At(",", cDate) + 1))
  Endif
Endif

If !Empty(cDate)
   nDay = 1
   nMonth = 1
   nYear = 1980

   If At(" ", cDate) <> 0
      nDay = Val(Left(cDate, At(" ", cDate) - 1))
      cDate = AllTrim(SubStr(cDate, At(" ", cDate)))

      If At(" ", cDate) <> 0
         cPart = Upper(Left(cDate, At(" ", cDate) - 1))
         cDate = AllTrim(SubStr(cDate, At(" ", cDate)))

         Do Case
            Case (cPart == Upper("Jan")) .OR. (cPart == Upper("Ene"))
                 nMonth = 1

            Case cPart == Upper("Feb")
                 nMonth = 2

            Case cPart == Upper("Mar")
                 nMonth = 3

            Case (cPart == Upper("Apr")) .OR. (cPart == Upper("Abr"))
                 nMonth = 4

            Case cPart == Upper("May")
                 nMonth = 5

            Case cPart == Upper("Jun")
                 nMonth = 6

            Case cPart == Upper("Jul")
                 nMonth = 7

            Case (cPart == Upper("Aug")) .OR. (cPart == Upper("Ago"))
                 nMonth = 8

            Case cPart == Upper("Sep")
                 nMonth = 9

            Case cPart == Upper("Oct")
                 nMonth = 10

            Case cPart == Upper("Nov")
                 nMonth = 11

            Case (cPart == Upper("Dec")) .OR. (cPart == Upper("Dic"))
                 nMonth = 12

         EndCase

         If At(" ", cDate) <> 0
            nYear = Val(Left(cDate, At(" ", cDate) - 1))
            If nYear < 100
               nYear = nYear + 2000
            Endif

            cDate = AllTrim(SubStr(cDate, At(" ", cDate)))

            Try
               dDate = Date(nYear, nMonth, nDay)
            Catch
            EndTry

            If Type("dDate") <> "D"
               dDate = CtoD("")
            Endif

            If At(" ", cDate) <> 0
               cTime = Left(cDate, At(" ", cDate) - 1)
               cDate = AllTrim(SubStr(cDate, At(" ", cDate)))
               nTimeZone = Val(cDate) / 100
            Endif
         Endif
      Endif
   Endif
   
Endif

this.SentDate = dDate
this.SentTime = cTime
this.SentTimeZone = nTimeZone


* Ajustar el time zone
oUtils = CreateObject("Utils")
nCurrentTimeZone = oUtils.GetTimeZoneOffset()

If (nCurrentTimeZone <> nTimeZone) .AND. (!Empty(dDate))
   nCurrentTimeZone = oUtils.GetTimeZoneOffsetMinutes()
   nDif = nCurrentTimeZone - nTimeZone * 60

   tTime = CtoT(DtoC(dDate) + " " + cTime)
   tTime = tTime + nDif * 60

   this.AdjustedDate = Date(Year(tTime), Month(tTime), Day(tTime))
   this.AdjustedTime = Right(TtoC(tTime), 8)
else
   this.AdjustedDate = this.SentDate
   this.AdjustedTime = this.SentTime
Endif

Return(dDate)



***
* Function RawMessage
***
Function RawMessage_Access()
Return(this.RawHeaders + ;
       Chr(13) + Chr(10) + Chr(13) + Chr(10) + ;
       this.RawBody)


EndDefine



**************************   MAILMESSAGE CLASS   ***************************

****************************************************************************
Define Class MailMessage as Custom

             Number = 0
             Size = 0

EndDefine



************************   MAILATTACHMENT CLASS   **************************

****************************************************************************
Define Class MailAttachment as Custom

             FileName = ""
             FileSize = 0
             Content = ""


***
* Function Init
***
Function Init(cContent)
this.Content = cContent
EndFunc



***
* Function Save
***
Function Save(cName)
StrToFile(this.Content, cName)
EndFunc

EndDefine



************************   MAILRESOURCE CLASS   ****************************

****************************************************************************
Define Class MailResource as Custom

             FileName = ""
             Content = ""
             ID = ""

EndDefine



************************   MAILRECEIVER CLASS   ****************************

****************************************************************************
Define Class MailReceiver as Custom

             ReceiverName = ""
             ReceiverEMail = ""

EndDefine
