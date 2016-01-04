
If .F.
   ?  "Creado por Pablo Pioli"
Endif


#include IFOX.H


**************************   SMTP CLASS   **********************************

****************************************************************************
Define Class SMTP as iFox of iFox.PRG OlePublic

       SenderName = ""
       SenderEMail = ""
       NotificationEMail = ""
       ReplyTo = ""

       Identification = "iFox"

       Subject = ""
       Priority = SMTP_PRIORITY_NORMAL
       IncludeImages = .F.
       IncludeMusic = .F.

       Body = ""
       BodyCharset = "iso-8859-1"
       BodyType = SMTP_BODY_STRING

       HTMLBody = ""
       HTMLBodyCharset = "iso-8859-1"
       HTMLBodyType = SMTP_BODY_STRING

       RTFBody = ""
       RTFBodyCharset = "iso-8859-1"
       RTFBodyType = SMTP_BODY_STRING

       BasePath = ""
       TimeOut = 10

       ErrorNumber = 0
       ExtraErrorInfo = ""
       ErrorDescription = ""

       SMTPErrorNumber = 0
       SMTPResponses = ""

       ValidAddresses = .T.
       NotifyDelivery = .F.
       FastGeneration = .F.

       AuthenticationMethod = 0
       UserName = ""
       Password = ""
       Pop3Server = ""
       EnableCRAMMD5 = .F.

       PartSize = 1024
       EOD = .F.
       Hidden CurrentPosition

       Hidden Recipientes[1, 2]
       Hidden CC[1, 2]
       Hidden BCC[1, 2]
       Hidden Attachments[1]
       Hidden Images[1, 2]
       Hidden ExtraHeaders[1]

       Hidden SocketsHandler
       Hidden ConnectionSocket
       Hidden Connected

       BodyData = ""
       PreprocessedMessage = ""
       Hidden Boundary1
       Hidden Boundary2
       Hidden Boundary3

       Hidden Utils
       Hidden FLLLoaded


***
* Function eInit
***
Protected Function eInit()


* Inicializar propiedades
this.SocketsHandler = CreateObject("Sockets")
this.ConnectionSocket = 0

this.Utils = CreateObject("Utils")
this.FLLLoaded = .F.

this.CurrentPosition = 1
this.BodyData = ""


EndFunc



***
* Function AddRecipient
***
Function AddRecipient(cName, cEMail)
Local lRes

If Type("cName") <> "C"
   cName = ""
Endif

If Type("cEMail") <> "C"
   cEMail = ""
Endif

If (!this.ValidAddresses) .OR. (this.ValidMailAddress(cEMail))
   Dimension this.Recipientes[ALen(this.Recipientes, 1) + 1, ALen(this.Recipientes, 2)]
   this.Recipientes[ALen(this.Recipientes, 1), 1] = cName
   this.Recipientes[ALen(this.Recipientes, 1), 2] = cEMail

   lRes = .T.
else
   lRes = .F.
Endif

Return(lRes)



***
* Function AddCC
***
Function AddCC(cName, cEMail)
Local lRes

If Type("cName") <> "C"
   cName = ""
Endif

If Type("cEMail") <> "C"
   cEMail = ""
Endif

If (!this.ValidAddresses) .OR. (this.ValidMailAddress(cEMail))
   Dimension this.CC[ALen(this.CC, 1) + 1, ALen(this.CC, 2)]
   this.CC[ALen(this.CC, 1), 1] = cName
   this.CC[ALen(this.CC, 1), 2] = cEMail

   lRes = .T.
else
   lRes = .F.
Endif

Return(lRes)



***
* Function AddBCC
***
Function AddBCC(cName, cEMail)
Local lRes

If Type("cName") <> "C"
   cName = ""
Endif

If Type("cEMail") <> "C"
   cEMail = ""
Endif

If (!this.ValidAddresses) .OR. (this.ValidMailAddress(cEMail))
   Dimension this.BCC[ALen(this.BCC, 1) + 1, ALen(this.BCC, 2)]
   this.BCC[ALen(this.BCC, 1), 1] = cName
   this.BCC[ALen(this.BCC, 1), 2] = cEMail

   lRes = .T.
else
   lRes = .F.
Endif

Return(lRes)



***
* Function AddAttachment
***
Function AddAttachment(cFile)
Local lRes

If File(cFile)
   Dimension this.Attachments[ALen(this.Attachments) + 1]
   this.Attachments[ALen(this.Attachments)] = cFile

   lRes = .T.
else
   lRes = .F.
Endif

Return(lRes)



***
* Function AddExtraHeader
***
Function AddExtraHeader(cHeader)

Dimension this.ExtraHeaders[ALen(this.ExtraHeaders) + 1]
this.ExtraHeaders[ALen(this.ExtraHeaders)] = cHeader

EndFunc


***
* Function ValidMailAddress
***
Function ValidMailAddress(cAddress)
Local lRes

lRes = .T.

If Empty(cAddress)
   lRes = .F.
Endif

If Len(cAddress) < 6
   lRes = .F.
Endif

If (At("@", cAddress) == 0) .OR. (At(".", cAddress) == 0)
   lRes = .F.
Endif

If At(" ", cAddress) <> 0
   lRes = .F.
Endif

Return(lRes)



***
* Function ErrorDescription_Access
***
Function ErrorDescription_Access
Local cRes

cRes = "Descripción no disponible"

If Type("this.ErrorNumber") == "N"

   Do Case
      Case this.ErrorNumber == 0
           cRes = "No se ha producido ningún error"

      Case this.ErrorNumber == 1
           cRes = "Error de conexión"

      Case this.ErrorNumber == 2
           cRes = "No se ha obtenido respuesta del servidor de correo"

      Case this.ErrorNumber == 3
           cRes = "La dirección de E-Mail del remitente es inválida (" + ;
                  this.ExtraErrorInfo + ")"

      Case this.ErrorNumber == 4
           cRes = "No se ha indicado ningún destinatario"

      Case this.ErrorNumber == 5
           cRes = "El servidor de correo ha respondido con un mensaje de error. (" + ;
                  this.ExtraErrorInfo + ")"

      Case this.ErrorNumber == 6
           cRes = "El servidor de correo ha rechazado la dirección de E-Mail del remitente. (" + ;
                  this.SenderEMail + ")"

      Case this.ErrorNumber == 7
           cRes = "El servidor de correo ha rechazado alguno de los destinatarios. (" + ;
                  this.ExtraErrorInfo + ")"

      Case this.ErrorNumber == 8
           cRes = "El servidor SMTP no soporta el método de autenticación seleccionado."

      Case this.ErrorNumber == 9
           cRes = "Error durante la autenticación. El servidor SMTP ha rechazado el nombre de usuario."

      Case this.ErrorNumber == 10
           cRes = "Error durante la autenticación. El servidor SMTP ha rechazado la clave de acceso."

      Case this.ErrorNumber == 11
           cRes = "Error durante la autenticación. El servidor SMTP ha rechazado el nombre de usuario o la clave de acceso."

      Case this.ErrorNumber == 999
           cRes = "Se ha llamado al método SendNextPart luego que todo el mensaje ha sido transmitido."

   EndCase

Endif

Return(cRes)



***
* Function ErrorText()
***
Function ErrorText(nLanguage)
Local cRes

If Type("nLanguage") <> "N"
   nLanguage = 1
Endif

If (nLanguage < 1) .OR. (nLanguage > 2)
   nLanguage = 1
Endif

If nLanguage == 1
   cRes = this.ErrorDescription
else
   cRes = "Description not available"

   If Type("this.ErrorNumber") == "N"

      Do Case
         Case this.ErrorNumber == 0
              cRes = "No error"

         Case this.ErrorNumber == 1
              cRes = "Connection error"

         Case this.ErrorNumber == 2
              cRes = "No response from mail server"

         Case this.ErrorNumber == 3
              cRes = "The sender address is invalid (" + ;
                     this.ExtraErrorInfo + ")"

         Case this.ErrorNumber == 4
              cRes = "No recipients have been set"

         Case this.ErrorNumber == 5
              cRes = "The mail server has responded with an error message. (" + ;
                     this.ExtraErrorInfo + ")"

         Case this.ErrorNumber == 6
              cRes = "The mail server has rejected the sender address. (" + ;
                     this.SenderEMail + ")"

         Case this.ErrorNumber == 7
              cRes = "The mail server has rejected some of the recipients address. (" + ;
                     this.ExtraErrorInfo + ")"

         Case this.ErrorNumber == 8
              cRes = "The mail server don't support the authentication method."

         Case this.ErrorNumber == 9
              cRes = "Error during authentication. The mail server has rejected the user name."

         Case this.ErrorNumber == 10
              cRes = "Error during authentication. The mail server has rejected the password."

         Case this.ErrorNumber == 11
              cRes = "Error during authentication. The mail server has rejected the user name or the password."

         Case this.ErrorNumber == 999
              cRes = "SendNextPart has been called after the whole message has been transmitted."

      EndCase

   Endif
Endif

Return(cRes)



***
* Function ErrorMessage
***
Function ErrorMessage(nLanguage)
Local cRes



cRes = "Error Desconocido"

If Type("this.ErrorNumber") == "N"

   Do Case
      Case this.ErrorNumber == 0
           cRes = "No se ha producido ningún error"

      Case this.ErrorNumber == 1
           cRes = "Error de conexión"

      Case this.ErrorNumber == 2
           cRes = "No se ha obtenido respuesta del servidor de correo"

      Case this.ErrorNumber == 3
           cRes = "La dirección de E-Mail del remitente es inválida (" + ;
                  this.ExtraErrorInfo + ")"

      Case this.ErrorNumber == 4
           cRes = "No se ha indicado ningún destinatario"

      Case this.ErrorNumber == 5
           cRes = "El servidor de correo ha respondido con un mensaje de error. (" + ;
                  this.ExtraErrorInfo + ")"

      Case this.ErrorNumber == 6
           cRes = "El servidor de correo ha rechazado la dirección de E-Mail del remitente. (" + ;
                  this.SenderEMail + ")"

      Case this.ErrorNumber == 7
           cRes = "El servidor de correo ha rechazado alguno de los destinatarios. (" + ;
                  this.ExtraErrorInfo + ")"

      Case this.ErrorNumber == 8
           cRes = "El servidor SMTP no soporta el método de autenticación seleccionado."

      Case this.ErrorNumber == 9
           cRes = "Error durante la autenticación. El servidor SMTP ha rechazado el nombre de usuario."

      Case this.ErrorNumber == 10
           cRes = "Error durante la autenticación. El servidor SMTP ha rechazado la clave de acceso."

      Case this.ErrorNumber == 11
           cRes = "Error durante la autenticación. El servidor SMTP ha rechazado el nombre de usuario o la clave de acceso."

      Case this.ErrorNumber == 999
           cRes = "Se ha llamado al método SendNextPart luego que todo el mensaje ha sido transmitido."

   EndCase

Endif

Return(cRes)



***
* Function NewMessage
***
Function NewMessage()

this.SenderName = ""
this.SenderEMail = ""
this.NotificationEMail = ""
this.ReplyTo = ""

this.Subject = ""
this.Priority = SMTP_PRIORITY_NORMAL
this.IncludeImages = .F.
this.IncludeMusic = .F.

this.Body = ""
this.BodyCharset = "iso-8859-1"
this.BodyType = SMTP_BODY_STRING

this.HTMLBody = ""
this.HTMLBodyCharset = "iso-8859-1"
this.HTMLBodyType = SMTP_BODY_STRING

this.RTFBody = ""
this.RTFBodyCharset = "iso-8859-1"
this.RTFBodyType = SMTP_BODY_STRING

Dimension this.Recipientes[1, 2]
Dimension this.CC[1, 2]
Dimension this.BCC[1, 2]
Dimension this.Attachments[1]
Dimension this.Images[1, 2]
Dimension this.ExtraHeaders[1]

this.ErrorNumber = 0
this.ExtraErrorInfo = ""
this.SMTPErrorNumber = 0
this.SMTPResponses = ""

this.BodyData = ""
this.PreprocessedMessage = ""
this.CurrentPosition = 1

EndFunc



***
* Function ConnectSend
***
Function ConnectSend(cServer, nPort)
Local lRes

lRes = .F.
If this.Connect(cServer, nPort)
   If this.Send()
      If this.Disconnect()
         lRes = .T.
      Endif
   Endif
Endif

Return(lRes)



***
* Function Connect
***
Function Connect(cServer, nPort)
Local lRes

If (Type("nPort") <> "N") .OR. (nPort == 0)
   nPort = 25
Endif

this.ConnectionSocket = this.SocketsHandler.Connect(cServer, nPort)

If this.ConnectionSocket <> 0
   lRes = .T.
else
   this.ErrorNumber = 1
   this.ExtraErrorInfo = ""
   this.SMTPErrorNumber = 0

   lRes = .F.
Endif

Return(lRes)



***
* Function Disconnect
***
Function Disconnect(lFastDisconnect)
Local lRes, nSeconds

lRes = .T.

If this.ConnectionSocket <> 0

   If lFastDisconnect
      nSeconds = 0
   else
      nSeconds = this.TimeOut
   Endif

   If this.SocketsHandler.SendReceive(this.ConnectionSocket, "QUIT" + NEW_LINE, nSeconds)
      this.SMTPResponses = this.SMTPResponses + this.SocketsHandler.Response

      If Left(this.SocketsHandler.Response, 3) <> "221"
         this.ErrorNumber = 5
         this.ExtraErrorInfo = this.SocketsHandler.Response
         this.SMTPErrorNumber = Val(Left(this.SocketsHandler.Response, 3))

         lRes = .F.
      Endif
   else
      this.ErrorNumber = 2
      this.ExtraErrorInfo = ""
      this.SMTPErrorNumber = 0

      lRes = .F.
   Endif

   this.SocketsHandler.Close(this.ConnectionSocket)
else
   this.ErrorNumber = 2
   this.ExtraErrorInfo = ""
   this.SMTPErrorNumber = 0

   lRes = .F.
Endif

this.Connected = .F.

Return(lRes)



***
* Function Send
***
Function Send()
Local lRes, lError

lRes = .F.

If this.StartPartialSending()

   lError = .F.
   Do While !this.EOD
      If !this.SendNextPart()
         lError = .T.
         Exit
      Endif
   Enddo

   If lError
      lRes = .F.
   else
      If this.EndPartialSending()
         lRes = .T.
      Endif
   Endif

Endif

Return(lRes)



***
* Function StartPartialSending
***
Function StartPartialSending()
Local lRes


this.SMTPResponses = ""
this.CurrentPosition = 1


If this.ConnectionSocket == 0
   this.ErrorNumber = 2
   this.ExtraErrorInfo = ""
   this.SMTPErrorNumber = 0

   Return(.F.)
Endif


If this.ValidAddresses
   If !this.ValidMailAddress(this.SenderEMail)
      this.ErrorNumber = 3
      this.ExtraErrorInfo = this.SenderEMail
      this.SMTPErrorNumber = 0

      Return(.F.)
   Endif
Endif


If (ALen(this.Recipientes, 1) <= 1) .AND. ;
   (ALen(this.CC, 1) <= 1) .AND. ;
   (ALen(this.BCC, 1) <= 1)

   this.ErrorNumber = 4
   this.ExtraErrorInfo = ""
   this.SMTPErrorNumber = 0

   Return(.F.)
Endif


lRes = .F.

If this.SendGreetings()
   If this.SendRecipients()
      If this.SendHeaders()

         If Empty(this.PreprocessedMessage)
            this.PrepareMail()
         else
            this.BodyData = this.PreprocessedMessage
         Endif
         lRes = .T.

      Endif
   Endif
Endif

Return(lRes)



***
* Function SendNextPart
***
Function SendNextPart(nRetries)
Local cData, lRes, nCont

cData = SubStr(this.BodyData, this.CurrentPosition, this.PartSize)

If Type("nRetries") <> "N"
   nRetries = 30
Endif

If Len(cData) > 0
   For nCont = 1 to nRetries
      If this.SocketsHandler.Send(this.ConnectionSocket, cData)
         lRes = .T.
         Exit
      else
         If nCont = nRetries
            this.ErrorNumber = 2
            this.ExtraErrorInfo = ""
            this.SMTPErrorNumber = 0

            lRes = .F.
            Exit
         else
            Win32Delay(1)
         Endif
      Endif
   Next
else
   this.ErrorNumber = 999
   this.ExtraErrorInfo = ""
   this.SMTPErrorNumber = 0

   lRes = .F.
Endif

this.CurrentPosition = this.CurrentPosition + Len(cData)

Return(lRes)



***
* Function EndPartialSending
***
Function EndPartialSending()
Local lRes

If this.SocketsHandler.SendReceive(this.ConnectionSocket, NEW_LINE + "." + NEW_LINE, this.TimeOut)
   this.SMTPResponses = this.SMTPResponses + this.SocketsHandler.Response

   If Left(this.SocketsHandler.Response, 3) == "250"
      lRes = .T.
   else
      this.ErrorNumber = 5
      this.ExtraErrorInfo = this.SocketsHandler.Response
      this.SMTPErrorNumber = Val(Left(this.SocketsHandler.Response, 3))

      lRes = .F.
   Endif
else
   this.ErrorNumber = 2
   this.ExtraErrorInfo = ""
   this.SMTPErrorNumber = 0

   lRes = .F.
Endif

Return(lRes)



***
* Function EOD_Access
***
Function EOD_Access()
Local lRes

If this.CurrentPosition > Len(this.BodyData)
   lRes = .T.
else
   lRes = .F.
Endif

Return(lRes)



***
* Function GetCurrentPosition()
***
Function GetCurrentPosition()
Local nRes

nRes = this.CurrentPosition

If nRes > Len(this.BodyData)
   nRes = Len(this.BodyData)
Endif

Return(nRes)



***
* Function GetMessageSize()
***
Function GetMessageSize()
Return(Len(this.BodyData))



***
* Function SendGreetings
***
Hidden Function SendGreetings()
Local lRes, cResponse, cData, oPOP3, tStart, cData

lRes = .T.


If !this.Connected

   If this.SocketsHandler.WaitFor(this.ConnectionSocket, "", this.TimeOut)
      this.SMTPResponses = this.SMTPResponses + this.SocketsHandler.Response

      If Left(this.SocketsHandler.Response, 3) <> "220"
         this.ErrorNumber = 5
         this.ExtraErrorInfo = this.SocketsHandler.Response
         this.SMTPErrorNumber = Val(Left(this.SocketsHandler.Response, 3))

         lRes = .F.
      Endif
   else
      this.ErrorNumber = 2
      this.ExtraErrorInfo = ""
      this.SMTPErrorNumber = 0

      lRes = .F.
   Endif


   If lRes
      If this.SocketsHandler.SendReceive(this.ConnectionSocket, "EHLO " + this.Identification + NEW_LINE, this.TimeOut)
         If Left(this.SocketsHandler.Response, 3) <> "250"
            this.ErrorNumber = 5
            this.ExtraErrorInfo = this.SocketsHandler.Response
            this.SMTPErrorNumber = Val(Left(this.SocketsHandler.Response, 3))

            lRes = .F.
         else
            tStart = DateTime()
            Do While .T.
               cData = this.SocketsHandler.Read(this.ConnectionSocket)
               If Empty(cData)
                  If DateTime() - tStart > 2
                     Exit
                  Endif
               else
                  this.SocketsHandler.Response = this.SocketsHandler.Response + cData

                  tStart = DateTime()
                  Do While .T.
                     cData = this.SocketsHandler.Read(this.ConnectionSocket)
                     If !Empty(cData)
                        this.SocketsHandler.Response = this.SocketsHandler.Response + cData
                     Endif
                     If DateTime() - tStart > 2
                        Exit
                     Endif
                  Enddo

                  Exit
               Endif
            Enddo
         Endif

         this.SMTPResponses = this.SMTPResponses + this.SocketsHandler.Response
      else
         this.ErrorNumber = 2
         this.ExtraErrorInfo = ""
         this.SMTPErrorNumber = 0

         lRes = .F.
      Endif
   Endif


   If lRes

      cResponse = Upper(this.SocketsHandler.Response)
      If At("250-AUTH", cResponse) == 0
         If At("250 AUTH", cResponse) == 0
            cResponse = ""
         else
            cResponse = SubStr(cResponse, At("250 AUTH ", Upper(cResponse)))
            cResponse = Left(cResponse, At(Chr(13), cResponse) - 1)
         Endif
      else
         cResponse = SubStr(cResponse, At("250-AUTH ", Upper(cResponse)))
         cResponse = Left(cResponse, At(Chr(13), cResponse) - 1)
      Endif

      Do Case
         Case this.AuthenticationMethod == 1
              oPOP3 = CreateObject("Pop3")
              oPop3.Username = this.UserName
              oPop3.Password = this.Password
              If oPop3.Connect(this.Pop3Server)
                 oPop3.Disconnect()
              else
                 this.ErrorNumber = 11
                 this.ExtraErrorInfo = ""
                 this.SMTPErrorNumber = 0

                 lRes = .F.
              Endif


         Case this.AuthenticationMethod == 2
              Do Case
              
              * To enable this section you need vfpencryption.fll
*!*                    Case (At("CRAM-MD5", cResponse) <> 0) .AND. (this.EnableCRAMMD5)
*!*                         If this.SocketsHandler.SendReceive(this.ConnectionSocket, "AUTH CRAM-MD5" + NEW_LINE, this.TimeOut)
*!*                            this.SMTPResponses = this.SMTPResponses + this.SocketsHandler.Response
*!*                            If Left(this.SocketsHandler.Response, 3) == "334"

*!*                               If !this.FLLLoaded
*!*                                  Set Library to (this.HomeDir + "vfpencryption.fll")
*!*                                  this.FLLLoaded = .T.
*!*                               Endif

*!*                               cData = StrConv(SubStr(this.SocketsHandler.Response, 5), 14)
*!*                               cData = this.UserName + " " + Lower(StrConv(hmac(cData, this.Password, 5), 15))
*!*                               cData = StrConv(cData, 13)

*!*                               If this.SocketsHandler.SendReceive(this.ConnectionSocket, cData + NEW_LINE, this.TimeOut)
*!*                                  this.SMTPResponses = this.SMTPResponses + this.SocketsHandler.Response

*!*                                  If Left(this.SocketsHandler.Response, 3) <> "235"
*!*                                     this.ErrorNumber = 10
*!*                                     this.ExtraErrorInfo = this.SocketsHandler.Response
*!*                                     this.SMTPErrorNumber = Val(Left(this.SocketsHandler.Response, 3))

*!*                                     lRes = .F.
*!*                                  Endif
*!*                               else
*!*                                  this.ErrorNumber = 2
*!*                                  this.ExtraErrorInfo = ""
*!*                                  this.SMTPErrorNumber = 0

*!*                                  lRes = .F.
*!*                               Endif

*!*                            else
*!*                               this.ErrorNumber = 9
*!*                               this.ExtraErrorInfo = this.SocketsHandler.Response
*!*                               this.SMTPErrorNumber = Val(Left(this.SocketsHandler.Response, 3))

*!*                               lRes = .F.
*!*                            Endif
*!*                         else
*!*                            this.ErrorNumber = 2
*!*                            this.ExtraErrorInfo = ""
*!*                            this.SMTPErrorNumber = 0

*!*                            lRes = .F.
*!*                         Endif


                 Case At("LOGIN", cResponse) <> 0
                      If this.SocketsHandler.SendReceive(this.ConnectionSocket, "AUTH LOGIN " + StrConv(this.UserName, 13) + NEW_LINE, this.TimeOut)
                         this.SMTPResponses = this.SMTPResponses + this.SocketsHandler.Response

                         If Left(this.SocketsHandler.Response, 3) == "334"
                            If this.SocketsHandler.SendReceive(this.ConnectionSocket, StrConv(this.Password, 13) + NEW_LINE, this.TimeOut)
                               this.SMTPResponses = this.SMTPResponses + this.SocketsHandler.Response

                               If Left(this.SocketsHandler.Response, 3) <> "235"
                                  this.ErrorNumber = 10
                                  this.ExtraErrorInfo = this.SocketsHandler.Response
                                  this.SMTPErrorNumber = Val(Left(this.SocketsHandler.Response, 3))

                                  lRes = .F.
                               Endif
                            else
                               this.ErrorNumber = 2
                               this.ExtraErrorInfo = ""
                               this.SMTPErrorNumber = 0

                               lRes = .F.
                            Endif
                         else
                            this.ErrorNumber = 9
                            this.ExtraErrorInfo = this.SocketsHandler.Response
                            this.SMTPErrorNumber = Val(Left(this.SocketsHandler.Response, 3))

                            lRes = .F.
                         Endif
                      else
                         this.ErrorNumber = 2
                         this.ExtraErrorInfo = ""
                         this.SMTPErrorNumber = 0

                         lRes = .F.
                      Endif


                 Case At("PLAIN", cResponse) <> 0
                      cData = this.UserName + Chr(0)
                      cData = cData + this.UserName + Chr(0)
                      cData = cData + this.Password
                      cData = StrConv(cData, 13)

                      If this.SocketsHandler.SendReceive(this.ConnectionSocket, "AUTH PLAIN " + cData + NEW_LINE, this.TimeOut)
                         this.SMTPResponses = this.SMTPResponses + this.SocketsHandler.Response

                         If Left(this.SocketsHandler.Response, 3) <> "235"
                            this.ErrorNumber = 11
                            this.ExtraErrorInfo = this.SocketsHandler.Response
                            this.SMTPErrorNumber = Val(Left(this.SocketsHandler.Response, 3))

                            lRes = .F.
                         Endif
                      else
                         this.ErrorNumber = 2
                         this.ExtraErrorInfo = ""
                         this.SMTPErrorNumber = 0

                         lRes = .F.
                      Endif


                 OtherWise
                      this.ErrorNumber = 8
                      this.ExtraErrorInfo = ""
                      this.SMTPErrorNumber = 0

                      lRes = .F.

              EndCase

      EndCase

   Endif


   this.Connected = .T.
Endif


If lRes
   If this.SocketsHandler.SendReceive(this.ConnectionSocket, "MAIL FROM:<" + this.SenderEMail + ">" + NEW_LINE, this.TimeOut)
      this.SMTPResponses = this.SMTPResponses + this.SocketsHandler.Response

      If Left(this.SocketsHandler.Response, 3) <> "250"
         this.ErrorNumber = 6
         this.ExtraErrorInfo = this.SocketsHandler.Response
         this.SMTPErrorNumber = Val(Left(this.SocketsHandler.Response, 3))

         lRes = .F.
      Endif
   else
      this.ErrorNumber = 2
      this.ExtraErrorInfo = ""
      this.SMTPErrorNumber = 0

      lRes = .F.
   Endif
Endif

Return(lRes)



***
* Function SendRecipients
***
Hidden Function SendRecipients()
Local cExtra, nCont


For nCont = 2 to ALen(this.Recipientes, 1)
    If this.NotifyDelivery
       cExtra = " NOTIFY=FAILURE,SUCCESS"
    else
       cExtra = ""
    Endif

    If this.SocketsHandler.SendReceive(this.ConnectionSocket, "RCPT TO:<" + this.Recipientes[nCont, 2] + ">" + cExtra + NEW_LINE, this.TimeOut)
       this.SMTPResponses = this.SMTPResponses + this.SocketsHandler.Response

       If Left(this.SocketsHandler.Response, 3) <> "250"
          this.ErrorNumber = 7
          this.ExtraErrorInfo = this.Recipientes[nCont, 2]
          this.SMTPErrorNumber = Val(Left(this.SocketsHandler.Response, 3))

          Return(.F.)
       Endif
    else
       this.ErrorNumber = 2
       this.ExtraErrorInfo = ""
       this.SMTPErrorNumber = 0

       Return(.F.)
    Endif
Next


For nCont = 2 to ALen(this.CC, 1)
    If this.SocketsHandler.SendReceive(this.ConnectionSocket, "RCPT TO:<" + this.CC[nCont, 2] + ">" + NEW_LINE, this.TimeOut)
       this.SMTPResponses = this.SMTPResponses + this.SocketsHandler.Response

       If Left(this.SocketsHandler.Response, 3) <> "250"
          this.ErrorNumber = 7
          this.ExtraErrorInfo = this.CC[nCont, 2]
          this.SMTPErrorNumber = Val(Left(this.SocketsHandler.Response, 3))

          Return(.F.)
       Endif
    else
       this.ErrorNumber = 2
       this.ExtraErrorInfo = ""
       this.SMTPErrorNumber = 0

       Return(.F.)
    Endif
Next


For nCont = 2 to ALen(this.BCC, 1)
    If this.SocketsHandler.SendReceive(this.ConnectionSocket, "RCPT TO:<" + this.BCC[nCont, 2] + ">" + NEW_LINE, this.TimeOut)
       this.SMTPResponses = this.SMTPResponses + this.SocketsHandler.Response

       If Left(this.SocketsHandler.Response, 3) <> "250"
          this.ErrorNumber = 7
          this.ExtraErrorInfo = this.BCC[nCont, 2]
          this.SMTPErrorNumber = Val(Left(this.SocketsHandler.Response, 3))

          Return(.F.)
       Endif
    else
       this.ErrorNumber = 2
       this.ExtraErrorInfo = ""
       this.SMTPErrorNumber = 0

       Return(.F.)
    Endif
Next


Return(.T.)



***
* Function SendHeaders
***
Hidden Function SendHeaders()
Local lRes

lRes = .T.

If this.SocketsHandler.SendReceive(this.ConnectionSocket, "DATA" + NEW_LINE, this.TimeOut)
   this.SMTPResponses = this.SMTPResponses + this.SocketsHandler.Response

   If Left(this.SocketsHandler.Response, 3) <> "354"
      this.ErrorNumber = 5
      this.ExtraErrorInfo = this.SocketsHandler.Response
      this.SMTPErrorNumber = Val(Left(this.SocketsHandler.Response, 3))

      lRes = .F.
   Endif

else
   this.ErrorNumber = 2
   this.ExtraErrorInfo = ""
   this.SMTPErrorNumber = 0

   lRes = .F.
Endif

Return(lRes)



***
* Function PrepareMail
***
Function PrepareMail()
Local cBody, cHTMLBody, cRTFBody, cGUID, cAux, cTag, nPos, cString, lEncode


* Leer el cuerpo, si es necesario
If this.BodyType == SMTP_BODY_FILE
   If File(this.Body)
      cBody = FiletoStr(this.Body)
   else
      cBody = ""
   Endif
else
   cBody = this.Body
Endif

If this.HTMLBodyType == SMTP_BODY_FILE
   If File(this.HTMLBody)
      cHTMLBody = FiletoStr(this.HTMLBody)
   else
      cHTMLBody = ""
   Endif
else
   cHTMLBody = this.HTMLBody
Endif

If this.RTFBodyType == SMTP_BODY_FILE
   If File(this.RTFBody)
      cRTFBody = FiletoStr(this.RTFBody)
   else
      cRTFBody = ""
   Endif
else
   cRTFBody = this.RTFBody
Endif


* Crear boundaries
cGUID = this.Utils.GenGUID()
cGUID = SubStr(cGUID, 2)
cGUID = Left(cGUID, Len(cGUID) - 1)
cGUID = StrTran(cGUID, "-", "_")
this.Boundary1 = "----=NextPart_" + cGUID + "=----"

cGUID = this.Utils.GenGUID()
cGUID = SubStr(cGUID, 2)
cGUID = Left(cGUID, Len(cGUID) - 1)
cGUID = StrTran(cGUID, "-", "_")
this.Boundary2 = "----=NextPart_" + cGUID + "=----"

cGUID = this.Utils.GenGUID()
cGUID = SubStr(cGUID, 2)
cGUID = Left(cGUID, Len(cGUID) - 1)
cGUID = StrTran(cGUID, "-", "_")
this.Boundary3 = "----=NextPart_" + cGUID + "=----"


* Agregar imagenes incrustadas
If this.IncludeImages

   cAux = cHTMLBody
   cHTMLBody = ""

   Do While Len(cAux) <> 0

      If At("<IMG", Upper(cAux)) == 0
         cHTMLBody = cHTMLBody + cAux
         cAux = ""
      else
         cHTMLBody = cHTMLBody + Left(cAux, At("<IMG", Upper(cAux)) - 1)
         cAux = SubStr(cAux, At("<IMG", Upper(cAux)))

         If At(">", cAux) == 0
            cHTMLBody = cHTMLBody + cAux
            cAux = ""
         else
            cTag = Left(cAux, At(">", cAux))
            cAux = SubStr(cAux, At(">", cAux) + 1)

            If At("SRC=", Upper(cTag)) <> 0
               cHTMLBody = cHTMLBody + Left(cTag, At("SRC=", Upper(cTag)) + 3)
               cTag = SubStr(cTag, At("SRC=", Upper(cTag)) + 4)

               If Left(cTag, 1) == '"'
                  cHTMLBody = cHTMLBody + '"'
                  cTag = SubStr(cTag, 2)

                  nPos = At('"', cTag)
                  If nPos == 0
                     cHTMLBody = cHTMLBody + cTag
                  else
                     cHTMLBody = cHTMLBody + this.AddFile(Left(cTag, At('"', cTag) - 1))
                     cHTMLBody = cHTMLBody + SubStr(cTag, At('"', cTag))
                  Endif
               else
                  nPos = At(" ", cTag)
                  If nPos == 0
                     cHTMLBody = cHTMLBody + cTag
                  else
                     cHTMLBody = cHTMLBody + this.AddFile(Left(cTag, At(" ", cTag) - 1))
                     cHTMLBody = cHTMLBody + SubStr(cTag, At(" ", cTag))
                  Endif
               Endif
            else
               cHTMLBody = cHTMLBody + cTag
            Endif

         Endif
      Endif

   Enddo

Endif


* Agregar musica
If this.IncludeMusic

   cAux = cHTMLBody
   cHTMLBody = ""

   Do While Len(cAux) <> 0

      If At("<BGSOUND", Upper(cAux)) == 0
         cHTMLBody = cHTMLBody + cAux
         cAux = ""
      else
         cHTMLBody = cHTMLBody + Left(cAux, At("<BGSOUND", Upper(cAux)) - 1)
         cAux = SubStr(cAux, At("<BGSOUND", Upper(cAux)))

         If At(">", cAux) == 0
            cHTMLBody = cHTMLBody + cAux
            cAux = ""
         else
            cTag = Left(cAux, At(">", cAux))
            cAux = SubStr(cAux, At(">", cAux) + 1)

            If At("SRC=", Upper(cTag)) <> 0
               cHTMLBody = cHTMLBody + Left(cTag, At("SRC=", Upper(cTag)) + 3)
               cTag = SubStr(cTag, At("SRC=", Upper(cTag)) + 4)

               If Left(cTag, 1) == '"'
                  cHTMLBody = cHTMLBody + '"'
                  cTag = SubStr(cTag, 2)

                  nPos = At('"', cTag)
                  If nPos == 0
                     cHTMLBody = cHTMLBody + cTag
                  else
                     cHTMLBody = cHTMLBody + this.AddFile(Left(cTag, At('"', cTag) - 1))
                     cHTMLBody = cHTMLBody + SubStr(cTag, At('"', cTag))
                  Endif
               else
                  nPos = At(" ", cTag)
                  If nPos == 0
                     cHTMLBody = cHTMLBody + cTag
                  else
                     cHTMLBody = cHTMLBody + this.AddFile(Left(cTag, At(" ", cTag) - 1))
                     cHTMLBody = cHTMLBody + SubStr(cTag, At(" ", cTag))
                  Endif
               Endif
            else
               cHTMLBody = cHTMLBody + cTag
            Endif

         Endif
      Endif

   Enddo

Endif


* Remitente
this.BodyData = "From: "

If this.Utils.NeedEncoding(this.SenderName)
   this.BodyData = this.BodyData + this.Utils.Encode(this.SenderName, 3, 6)
else
   this.BodyData = this.BodyData + '"' + this.SenderName + '"'
Endif

this.BodyData = this.BodyData + " <" + this.SenderEMail + ">" + NEW_LINE


* Destinatarios
If (ALen(this.Recipientes, 1) <= 1) .AND. (ALen(this.CC, 1) <= 1)
   this.BodyData = this.BodyData + "To: <Undisclosed-Recipient:;>" + NEW_LINE
else

   cString = ""
   For nCont = 2 to ALen(this.Recipientes, 1)
       If Len(cString) <> 0
          cString = cString + "," + NEW_LINE + " "
       Endif

       If this.Utils.NeedEncoding(this.Recipientes[nCont, 1])
          cString = cString + this.Utils.Encode(this.Recipientes[nCont, 1], 3, 4)
       else
          cString = cString + '"' + this.Recipientes[nCont, 1] + '"'
       Endif
       
       cString = cString + " <" + this.Recipientes[nCont, 2] + ">"
   Next
   If !Empty(cString)
      this.BodyData = this.BodyData + "To: " + cString + NEW_LINE
   Endif


   cString = ""
   For nCont = 2 to ALen(this.CC, 1)
       If Len(cString) <> 0
          cString = cString + "," + NEW_LINE + " "
       Endif

       If this.Utils.NeedEncoding(this.CC[nCont, 1])
          cString = cString + this.Utils.Encode(this.CC[nCont, 1], 3, 4)
       else
          cString = cString + '"' + this.CC[nCont, 1] + '"'
       Endif

       cString = cString + " <" + this.CC[nCont, 2] + ">"
   Next
   If !Empty(cString)
      this.BodyData = this.BodyData + "CC: " + cString + NEW_LINE
   Endif

Endif


* Reply-To
If !Empty(this.ReplyTo)
   this.BodyData = this.BodyData + "Reply-To: " + this.ReplyTo + NEW_LINE
Endif


* Asunto
this.BodyData = this.BodyData + "Subject: "

If this.Utils.NeedEncoding(this.Subject)
   this.BodyData = this.BodyData + this.Utils.Encode(this.Subject, 3, 9)
else
   this.BodyData = this.BodyData + this.Subject
Endif

this.BodyData = this.BodyData + NEW_LINE


* Version MIME
this.BodyData = this.BodyData + "MIME-Version: 1.0" + NEW_LINE


* Multipart
If (Len(cHTMLBody) <> 0) .OR. (Len(cRTFBody) <> 0) .OR. (ALen(this.Attachments) > 1)
   If ALen(this.Attachments) > 1
      this.BodyData = this.BodyData + "Content-Type: multipart/mixed;" + NEW_LINE
   else
      If (this.IncludeImages) .OR. (this.IncludeMusic)
         this.BodyData = this.BodyData + "Content-Type: multipart/related;" + NEW_LINE
      else
         this.BodyData = this.BodyData + "Content-Type: multipart/alternative;" + NEW_LINE
      Endif
   Endif

   this.BodyData = this.BodyData + ' boundary="' + this.Boundary1 + '"' + NEW_LINE
else
   this.BodyData = this.BodyData + 'Content-Type: text/plain;charset="' + this.BodyCharset + '"' + NEW_LINE
   this.BodyData = this.BodyData + "Content-Transfer-Encoding: "

   lEncode = this.Utils.NeedEncoding(cBody)
   If lEncode
      this.BodyData = this.BodyData + "quoted-printable"
   else
      this.BodyData = this.BodyData + "7bit"
   Endif
   
   this.BodyData = this.BodyData + NEW_LINE
Endif


* La fecha
this.BodyData = this.BodyData + "Date: " + this.GetSendDate() + NEW_LINE


* Un poco de promocion
this.BodyData = this.BodyData + "X-Mailer: iFox (www.coliseosoftware.com.ar/ifox)" + NEW_LINE


* Prioridad
this.BodyData = this.BodyData + "X-Priority: " + LTrim(Str(this.Priority, 3, 0)) + NEW_LINE


* Solicitud de notificacion de reception
If !Empty(this.NotificationEMail)
   this.BodyData = this.BodyData + "Disposition-Notification-To: " + this.NotificationEMail + NEW_LINE
Endif


* Encabezados extras
For nCont = 2 to ALen(this.ExtraHeaders)
   this.BodyData = this.BodyData + this.ExtraHeaders[nCont] + NEW_LINE
Next


* Cuerpo
this.BodyData = this.BodyData + NEW_LINE


If (Len(cHTMLBody) <> 0) .OR. (Len(cRTFBody) <> 0) .OR. (ALen(this.Attachments) > 1)
   this.BodyData = this.BodyData + "This is a multi-part message in MIME format." + NEW_LINE
   this.BodyData = this.BodyData + NEW_LINE
   this.BodyData = this.BodyData + "--" + this.Boundary1


   * Body
   If (Len(cHTMLBody) == 0) .AND. (Len(cRTFBody) == 0)

      * Solo Texto
      this.AddBody(cBody, this.Boundary1)

   else
      If ALen(this.Attachments) > 1
         If (this.IncludeImages) .OR. (this.IncludeMusic)

            * Texto, HTML, e imagenes
            this.BodyData = this.BodyData + NEW_LINE
            this.BodyData = this.BodyData + "Content-Type: multipart/related;" + NEW_LINE
            this.BodyData = this.BodyData + ' boundary="' + this.Boundary2 + '"' + NEW_LINE + NEW_LINE + NEW_LINE
            this.BodyData = this.BodyData + "--" + this.Boundary2 + NEW_LINE

            this.BodyData = this.BodyData + "Content-Type: multipart/alternative;" + NEW_LINE
            this.BodyData = this.BodyData + ' boundary="' + this.Boundary3 + '"' + NEW_LINE + NEW_LINE + NEW_LINE
            this.BodyData = this.BodyData + "--" + this.Boundary3

            this.AddBody(cBody, this.Boundary3)
            this.AddHTMLBody(cHTMLBody, this.Boundary3 + "--")
            this.AddRTFBody(cRTFBody, this.Boundary3 + "--")

            this.BodyData = this.BodyData + NEW_LINE + NEW_LINE
            this.BodyData = this.BodyData + "--" + this.Boundary2

            If ALen(this.Images, 1) > 1
               this.AddImages(this.Boundary2)
               this.BodyData = this.BodyData + "--"
            else
               this.BodyData = this.BodyData + NEW_LINE + "--" + this.Boundary2 + "--"
            Endif

            this.BodyData = this.BodyData + NEW_LINE + NEW_LINE
            this.BodyData = this.BodyData + "--" + this.Boundary1

         else

            * Texto y HTML
            this.BodyData = this.BodyData + NEW_LINE
            this.BodyData = this.BodyData + "Content-Type: multipart/alternative;" + NEW_LINE
            this.BodyData = this.BodyData + ' boundary="' + this.Boundary2 + '"' + NEW_LINE + NEW_LINE + NEW_LINE
            this.BodyData = this.BodyData + "--" + this.Boundary2 + NEW_LINE

            this.AddBody(cBody, this.Boundary2)
            this.AddHTMLBody(cHTMLBody, this.Boundary2 + "--")
            this.AddRTFBody(cRTFBody, this.Boundary2 + "--")

            this.BodyData = this.BodyData + NEW_LINE + NEW_LINE
            this.BodyData = this.BodyData + "--" + this.Boundary1

            this.AddImages(this.Boundary1)

         Endif
      else
         If (this.IncludeImages) .OR. (this.IncludeMusic)

            * Texto, HTML, e imagenes
            this.BodyData = this.BodyData + NEW_LINE
            this.BodyData = this.BodyData + "Content-Type: multipart/alternative;" + NEW_LINE
            this.BodyData = this.BodyData + ' boundary="' + this.Boundary2 + '"' + NEW_LINE + NEW_LINE
            this.BodyData = this.BodyData + "--" + this.Boundary2

            this.AddBody(cBody, this.Boundary2)
            this.AddHTMLBody(cHTMLBody, this.Boundary2 + "--")
            this.AddRTFBody(cRTFBody, this.Boundary2 + "--")

            this.BodyData = this.BodyData + NEW_LINE + NEW_LINE
            this.BodyData = this.BodyData + "--" + this.Boundary1

            this.AddImages(this.Boundary1)

         else

            * Texto y HTML
            this.AddBody(cBody, this.Boundary1)
            this.AddHTMLBody(cHTMLBody, this.Boundary1)
            this.AddRTFBody(cRTFBody, this.Boundary1)

         Endif
      Endif
   Endif


   * Adjuntos
   this.AddAttachments()


   * Terminador de mensaje
   this.BodyData = this.BodyData + "--" + NEW_LINE

else
   If lEncode
      this.BodyData = this.BodyData + this.Utils.Encode(cBody, 1)
   else
      this.BodyData = this.BodyData + cBody
   Endif
Endif


* Para evitar errores
this.BodyData = StrTran(this.BodyData, NEW_LINE + "." + NEW_LINE, NEW_LINE + ". " + NEW_LINE)


EndFunc



***
* Function AddFile
***
Hidden Function AddFile(cFile)
Local cRes, nCont, cGUID

cRes = ""

If (Left(cFile, 2) <> "\\") .AND. (SubStr(cFile, 2, 1) <> ":")
   If Empty(this.BasePath)
      cFile = AddBS(JustPath(this.HTMLBody)) + cFile
   else
      cFile = AddBS(this.BasePath) + cFile
   Endif
Endif

If File(cFile)

   For nCont = 2 to ALen(this.Images, 1)
       If Upper(this.Images[nCont, 1]) == Upper(cFile)
          Exit
       Endif
   Next

   If nCont == ALen(this.Images, 1) + 1
      cGUID = this.Utils.GenGUID()
      cGUID = SubStr(cGUID, 2)
      cGUID = Left(cGUID, Len(cGUID) - 1)
      cGUID = StrTran(cGUID, "-", "_")

      Dimension this.Images[nCont, ALen(this.Images, 2)]
      this.Images[nCont, 1] = cFile
      this.Images[nCont, 2] = cGUID
   Endif

   cRes = "cid:" + this.Images[nCont, 2]
Endif

Return(cRes)



***
* Function AddBody
***
Hidden Function AddBody(cBody, cBoundary)
Local lEncode

this.BodyData = this.BodyData + NEW_LINE
this.BodyData = this.BodyData + 'Content-Type: text/plain;charset="' + this.BodyCharset + '"' + NEW_LINE
this.BodyData = this.BodyData + "Content-Transfer-Encoding: "

lEncode = this.Utils.NeedEncoding(cBody)

If lEncode
   this.BodyData = this.BodyData + "quoted-printable"
else
   this.BodyData = this.BodyData + "7bit"
Endif

this.BodyData = this.BodyData + NEW_LINE + NEW_LINE

If lEncode
   this.BodyData = this.BodyData + this.Utils.Encode(cBody, 1)
else
   this.BodyData = this.BodyData + cBody
Endif

this.BodyData = this.BodyData + NEW_LINE + NEW_LINE
this.BodyData = this.BodyData + "--" + cBoundary

EndFunc



***
* Function AddHTMLBody
***
Hidden Function AddHTMLBody(cHTMLBody, cBoundary)

If !Empty(cHTMLBody)
   this.BodyData = this.BodyData + NEW_LINE
   this.BodyData = this.BodyData + 'Content-Type: text/html;charset="' + this.HTMLBodyCharset + '"' + NEW_LINE
   this.BodyData = this.BodyData + "Content-Transfer-Encoding: quoted-printable" + NEW_LINE
   this.BodyData = this.BodyData + NEW_LINE
   this.BodyData = this.BodyData + this.Utils.Encode(cHTMLBody, 1)
   this.BodyData = this.BodyData + NEW_LINE + NEW_LINE
   this.BodyData = this.BodyData + "--" + cBoundary
Endif

EndFunc



***
* Function AddRTFBody
***
Hidden Function AddRTFBody(cRTFBody, cBoundary)

If !Empty(cRTFBody)
   this.BodyData = this.BodyData + NEW_LINE
   this.BodyData = this.BodyData + 'Content-Type: text/rtf;charset="' + this.RTFBodyCharset + '"' + NEW_LINE
   this.BodyData = this.BodyData + "Content-Transfer-Encoding: quoted-printable" + NEW_LINE
   this.BodyData = this.BodyData + NEW_LINE
   this.BodyData = this.BodyData + this.Utils.Encode(cRTFBody, 1)
   this.BodyData = this.BodyData + NEW_LINE + NEW_LINE
   this.BodyData = this.BodyData + "--" + cBoundary
Endif

EndFunc



***
* Function AddImages
***
Hidden Function AddImages(cBoundary)
Local nCont, cData

For nCont = 2 to ALen(this.Images, 1)
    this.BodyData = this.BodyData + NEW_LINE
    this.BodyData = this.BodyData + "Content-Type: " + this.Utils.GetType(JustExt(this.Images[nCont, 1])) + ";" + ;
                                    'Name="' + JustFName(this.Images[nCont, 1]) + '"' + NEW_LINE
    this.BodyData = this.BodyData + "Content-Transfer-Encoding: base64" + NEW_LINE
    this.BodyData = this.BodyData + "Content-ID: <" + ;
                                    this.Images[nCont, 2] + ">" + NEW_LINE
    this.BodyData = this.BodyData + NEW_LINE
       
    cData = StrConv(FiletoStr(this.Images[nCont, 1]), 13)

    If this.FastGeneration
       this.BodyData = this.BodyData + cData + NEW_LINE
    else
       Do While !Empty(cData)

          If Len(cData) > 76
             this.BodyData = this.BodyData + Left(cData, 76) + NEW_LINE
             cData = SubStr(cData, 77)
          else
             this.BodyData = this.BodyData + cData + NEW_LINE
             cData = ""
          Endif

        Enddo
     Endif

    this.BodyData = this.BodyData + NEW_LINE
    this.BodyData = this.BodyData + "--" + cBoundary
Next

EndFunc



***
* Function AddAttachments
***
Hidden Function AddAttachments()
Local nCont, cData

For nCont = 2 to ALen(this.Attachments)
    this.BodyData = this.BodyData + NEW_LINE
    this.BodyData = this.BodyData + "Content-Type: " + this.Utils.GetType(JustExt(this.Attachments[nCont])) + ";" + ;
                                    'Name="' + JustFName(this.Attachments[nCont]) + '"' + NEW_LINE
    this.BodyData = this.BodyData + "Content-Transfer-Encoding: base64" + NEW_LINE
    this.BodyData = this.BodyData + "Content-Disposition: attachment;" + ;
                                    'filename="' + JustFName(this.Attachments[nCont]) + '"' + NEW_LINE
    this.BodyData = this.BodyData + NEW_LINE

    cData = StrConv(FiletoStr(this.Attachments[nCont]), 13)

    If this.FastGeneration
       Do While !Empty(cData)
          If Len(cData) > 1024 * 1024
             this.BodyData = this.BodyData + Left(cData, 1024 * 1024) + NEW_LINE
             cData = SubStr(cData, 1024 * 1024 + 1)
          else
             this.BodyData = this.BodyData + cData + NEW_LINE
             cData = ""
          Endif
       Enddo
    else
       Do While !Empty(cData)
          If Len(cData) > 76
             this.BodyData = this.BodyData + Left(cData, 76) + NEW_LINE
             cData = SubStr(cData, 77)
          else
             this.BodyData = this.BodyData + cData + NEW_LINE
             cData = ""
          Endif
       Enddo
    Endif

    this.BodyData = this.BodyData + NEW_LINE
    this.BodyData = this.BodyData + "--" + this.Boundary1
Next

EndFunc



***
* Function GetSendDate
***
Hidden Function GetSendDate()
Local cRes, dDate, dDay, nMonth, nZone

cRes = ""
dDate = Date()

dDay = DoW(dDate, 2)

Do Case
   Case dDay == 1
        cRes = cRes + "Mon"

   Case dDay == 2
        cRes = cRes + "Tue"

   Case dDay == 3
        cRes = cRes + "Wed"

   Case dDay == 4
        cRes = cRes + "Thu"

   Case dDay == 5
        cRes = cRes + "Fri"

   Case dDay == 6
        cRes = cRes + "Sat"

   Case dDay == 7
        cRes = cRes + "Sun"

EndCase


cRes = cRes + ", "
cRes = cRes + PadL(Day(dDate), 2, "0") + " "

nMonth = Month(dDate)

Do Case
   Case nMonth == 1
        cRes = cRes + "Jan"

   Case nMonth == 2
        cRes = cRes + "Feb"

   Case nMonth == 3
        cRes = cRes + "Mar"

   Case nMonth == 4
        cRes = cRes + "Apr"

   Case nMonth == 5
        cRes = cRes + "May"

   Case nMonth == 6
        cRes = cRes + "Jun"

   Case nMonth == 7
        cRes = cRes + "Jul"

   Case nMonth == 8
        cRes = cRes + "Aug"

   Case nMonth == 9
        cRes = cRes + "Sep"

   Case nMonth == 10
        cRes = cRes + "Oct"

   Case nMonth == 11
        cRes = cRes + "Nov"

   Case nMonth == 12
        cRes = cRes + "Dec"

EndCase

cRes = cRes + " " + PadL(Year(dDate), 4, "0") + " "
cRes = cRes + Time() + " "


nZone = this.Utils.GetTimeZoneOffset()

Do Case
   Case nZone < 0
        cRes = cRes + "-"

   Case nZone > 0
        cRes = cRes + "+"
EndCase

cRes = cRes + PadL(Int(Abs(nZone)), 2, "0") + "00"

Return(cRes)


EndDefine
