
If .F.
   ?  "Creado por Pablo Pioli"
Endif


#include IFOX.H


#define INTERNET_DEFAULT_HTTP_PORT          80
#define INTERNET_DEFAULT_HTTPS_PORT        443

#define INTERNET_OPEN_TYPE_WINDOWSCONFIG     0
#define INTERNET_OPEN_TYPE_DIRECT            1
#define INTERNET_OPEN_TYPE_PROXY             3

#define INTERNET_SERVICE_HTTP                3
#define INTERNET_FLAG_RELOAD        2147483648

#define INTERNET_OPTION_PROXY_USERNAME      43
#define INTERNET_OPTION_PROXY_PASSWORD      44

#define INTERNET_FLAG_SECURE        0x00800000

#define INTERNET_OPTION_SECURITY_FLAGS         31
#define SECURITY_FLAG_IGNORE_CERT_DATE_INVALID 0x00002000
#define SECURITY_FLAG_IGNORE_CERT_CN_INVALID   0x00001000
#define SECURITY_FLAG_IGNORE_UNKNOWN_CA        0x00000100
#define SECURITY_FLAG_IGNORE_WRONG_USAGE       0x00000200
#define SECURITY_FLAG_IGNORE_REVOCATION        0x00000080



**************************   HTTP CLASS   **********************************

****************************************************************************
Define Class HTTP as iFox of iFox.PRG OlePublic

       Protected FieldCount
       Dimension PostFields[1, 3]

       Protected Upload
       Protected AsyncInetHandle
       Protected AsyncURLHandle
       Protected AsyncHTTPSession
       Protected AsyncHTTPResult
       Protected AsyncData

       ConnectionType = INTERNET_OPEN_TYPE_WINDOWSCONFIG
       ProxyServer = ""
       ProxyUserName = ""
       ProxyPassword = ""

       TransferSize = 0
       CompletedTransfer = 0
       EOT = .F.
       BatchSize = 16384
       Response = ""

       StatusCode = 0
       ErrorNumber = 0

       IgnoreSSLErrors = .F.

       PostHeaders = ""
       PostData = ""


***
* Function eInit
***
Function eInit()


* Declare WinInet Functions
Declare Integer InternetOpen in WinInet.DLL String, Integer, String, String, Integer
Declare Integer InternetCloseHandle in WinInet.DLL Integer
Declare Integer InternetOpenUrl in WinInet.DLL Integer, String, String, Integer, Integer, Integer
Declare Integer InternetReadFile in WinInet.DLL Integer, String @cBuffer, Integer nBuffer, Integer @nSizeRead
Declare Integer InternetWriteFile In WinInet.DLL Integer hFile, String @cBuffer, Integer lNumBytesToWrite, Integer @dwNumberOfBytesWritten
Declare Integer InternetConnect in WinInet.DLL Integer hIPHandle, String lpzServer, Integer dwPort, String lpzUserName, String lpzPassword, Integer dwServiceFlags, Integer dwReserved, Integer dwReserved
Declare Integer HttpOpenRequest In WinInet.DLL Integer hHTTPHandle, String lpzReqMethod, String lpzPage, String lpzVersion, String lpzReferer, String lpzAcceptTypes, Integer dwFlags, Integer dwContext
Declare Integer HttpAddRequestHeaders In WinInet.DLL Integer hHTTPHandle, String lpzHeaders, Integer cbHeaders, Integer Options
Declare Integer HttpSendRequest In WinInet.DLL Integer hHTTPHandle, String lpzHeaders, Integer cbHeaders, String lpzPost, Integer cbPost
Declare Integer HttpSendRequestEx In WinInet.DLL Integer hHTTPHandle, String BuffersIn, Integer BuffersOut, Integer dwFlags, Integer dwContext
Declare Integer HttpEndRequest In WinInet.DLL Integer hHTTPHandle, Integer BuffersOut, Integer dwFlags, Integer dwContext
Declare Integer HttpQueryInfo In WinInet.DLL Integer hHTTPHandle, Long dwInfoLevel, String @lpvBuffer, Long @lpdwBufferLength, Long @lpdwIndex
Declare Integer InternetQueryOption In WinInet.DLL Integer hInternet, Integer lOption, String @sBuffer, Long lBufferLength
        
Declare Integer GetLastError in Win32API


* Inicializar propiedades
this.FieldCount = 0
this.Upload = .F.

EndFunc



***
* Function Get
***
Function Get(cURL)
Local nInetHandle, nURLHandle, cHeader, cStatusCode, nResult, cBuffer, nSize, cRes


* Inicializar Variables
this.ErrorNumber = 0
this.StatusCode = 0


* Conectarse a los servicios de Internet
nInetHandle = InternetOpen("iFox", this.ConnectionType, this.ProxyServer, "", 0)

If nInetHandle == 0
   this.ErrorNumber = GetLastError()
   Return("")
Endif


* Abrir la URL
cHeader = "Accept: */*" + Chr(13) + Chr(13)

nURLHandle = InternetOpenUrl(nInetHandle, cURL, cHeader, Len(cHeader), ;
                             INTERNET_FLAG_RELOAD, 0)

If nURLHandle == 0
   this.ErrorNumber = GetLastError()
   InternetCloseHandle(nInetHandle)
   Return("")
Endif


* Establecer opciones de autenticacion con el proxy
If !this.SetProxyAuthentication(nURLHandle)
   InternetCloseHandle(nURLHandle)
   InternetCloseHandle(nInetHandle)
   Return("")
Endif


* Leer respuesta
cRes = ""

Do While .T.

   cBuffer = Space(65536)
   nSize = 65536

   nResult = InternetReadFile(nURLHandle, @cBuffer, Len(cBuffer), @nSize)

   If nResult == 0
      Exit
   else
      If nSize == 0
         Exit
      else
         cBuffer = IIF(nSize > 1, SubStr(cBuffer, 1, nSize), "")
         cRes = cRes + cBuffer
      Endif
   Endif

Enddo


cStatusCode = Space(10)
HttpQueryInfo(nURLHandle, 19, @cStatusCode, Len(cStatusCode), 0)
this.StatusCode = Val(cStatusCode)


InternetCloseHandle(nURLHandle)
InternetCloseHandle(nInetHandle)


If nResult == 0
   this.ErrorNumber = GetLastError()
   Return("")
Endif

Return(cRes)



***
* Function SetProxyAuthentication
***
Function SetProxyAuthentication(nURLHandle)
Local cProxyUserName, cProxyPassword

cProxyUserName = this.ProxyUserName
cProxyPassword = this.ProxyPassword

If (!Empty(cProxyUserName))
   Declare Integer InternetSetOption In WinInet.DLL Integer hInternet, Integer dwFlags, String @dwValue, Long cbSize

   cProxyUserName = cProxyUserName + Chr(0)
   If InternetSetOption(nURLHandle, INTERNET_OPTION_PROXY_USERNAME, @cProxyUserName, Len(cProxyUserName)) <> 1
      this.ErrorNumber = GetLastError()
      Return(.F.)
   Endif

   If !Empty(cProxyPassword)
      cProxyPassword = cProxyPassword + Chr(0)
      If InternetSetOption(nURLHandle, INTERNET_OPTION_PROXY_PASSWORD, @cProxyPassword, Len(cProxyPassword)) <> 1
         this.ErrorNumber = GetLastError()
         Return(.F.)
      Endif
   Endif
Endif

Return(.T.)



***
* Function StartAsyncGet
***
Function StartAsyncGet(cURL)
Local cHeader

* Inicializar Variables
this.ErrorNumber = 0
this.StatusCode = 0

this.AsyncURLHandle = 0
this.AsyncInetHandle = 0



* Conectarse a los servicios de Internet
this.AsyncInetHandle = InternetOpen("iFox", this.ConnectionType, this.ProxyServer, "", 0)

If this.AsyncInetHandle == 0
   this.ErrorNumber = GetLastError()
   Return(.F.)
Endif


* Abrir la URL
cHeader = "Accept: */*" + Chr(13) + Chr(13)

this.AsyncURLHandle = InternetOpenUrl(this.AsyncInetHandle, cURL, cHeader, Len(cHeader), ;
                                      INTERNET_FLAG_RELOAD, 0)

If this.AsyncURLHandle == 0
   this.ErrorNumber = GetLastError()
   InternetCloseHandle(this.AsyncInetHandle)
   Return(.F.)
Endif


* Establecer opciones de autenticacion con el proxy
If !this.SetProxyAuthentication(this.AsyncURLHandle)
   InternetCloseHandle(this.AsyncURLHandle)
   InternetCloseHandle(this.AsyncInetHandle)
   Return(.F.)
Endif


this.TransferSize = 0
this.CompletedTransfer = 0

this.Response = ""
this.EOT = .F.

Return(.T.)



***
* Function ContinueAsyncGet
***
Function ContinueAsyncGet()
Local cSize, cBuffer, nSize, nResult

If this.TransferSize == 0
   cSize = Space(10)
   If HttpQueryInfo(this.AsyncURLHandle, 5, @cSize, Len(cSize), 0) <> 0
      this.TransferSize = Val(cSize)
   Endif
Endif

If this.EOT
   Return(.F.)
else
   cBuffer = Space(this.BatchSize)
   nSize = this.BatchSize
   nResult = InternetReadFile(this.AsyncURLHandle, @cBuffer, Len(cBuffer), @nSize)

   If nResult == 0
      Return(.F.)
   else
      If nSize == 0
         this.Response = ""
         this.EOT = .T.
         Return(.T.)
      else
         cBuffer = IIF(nSize > 1, SubStr(cBuffer, 1, nSize), "")
         this.CompletedTransfer = this.CompletedTransfer + nSize
         this.Response = cBuffer
      Endif
   Endif
Endif

Return(.T.)



***
* Function EndAsyncGet
***
Function EndAsyncGet()
Local cStatusCode

cStatusCode = Space(10)
HttpQueryInfo(this.AsyncURLHandle, 19, @cStatusCode, Len(cStatusCode), 0)
this.StatusCode = Val(cStatusCode)

InternetCloseHandle(this.AsyncURLHandle)
InternetCloseHandle(this.AsyncInetHandle)

EndFunc



***
* Function ClearPostFields
***
Function ClearPostFields()

this.FieldCount = 0
Dimension this.PostFields[1, 4]

this.Upload = .F.

EndFunc



***
* Function AddPostField
***
Function AddPostField(cField, cValue)

this.FieldCount = this.FieldCount + 1

Dimension this.PostFields[this.FieldCount, 4]
this.PostFields[this.FieldCount, 1] = 1
this.PostFields[this.FieldCount, 2] = cField
this.PostFields[this.FieldCount, 3] = cValue
this.PostFields[this.FieldCount, 4] = ""

EndFunc



***
* Function AddPostFile
***
Function AddPostFile(cField, cFile)

If !File(cFile)
   Return(.F.)
Endif

this.FieldCount = this.FieldCount + 1

Dimension this.PostFields[this.FieldCount, 4]
this.PostFields[this.FieldCount, 1] = 2
this.PostFields[this.FieldCount, 2] = cField
this.PostFields[this.FieldCount, 3] = cFile
this.PostFields[this.FieldCount, 4] = ""

this.Upload = .T.

Return(.T.)



***
* Function AddPostFieldasFile
***
Function AddPostFieldasFile(cField, cValue, cFile)

this.FieldCount = this.FieldCount + 1

Dimension this.PostFields[this.FieldCount, 4]
this.PostFields[this.FieldCount, 1] = 3
this.PostFields[this.FieldCount, 2] = cField
this.PostFields[this.FieldCount, 3] = cValue
this.PostFields[this.FieldCount, 4] = cFile

this.Upload = .T.

Return(.T.)



***
* Function Post
***
Function Post(cServer, cURL, cUsername, cPassword, nPort, cVerb)
Local nPort, nInetHandle, nHTTPSession, nHTTPResult
Local cPostBuffer, cHeaders, cReadBuffer, nRetVal, nBytesRead, cRes
Local cBoundary, lFound, cStatusCode, nFlags


* Inicializar Variables
this.ErrorNumber = 0
this.StatusCode = 0


If Type("cUserName") <> "C"
   cUserName = ""
Endif

If Type("cPassword") <> "C"
   cPassword = ""
Endif

If Type("nPort") <> "N"
   nPort = INTERNET_DEFAULT_HTTP_PORT
Endif

If Type("cVerb") <> "N"
   cVerb = "POST"
Endif


* Inicializar los servicios de Internet
nInetHandle = InternetOpen("iFox", this.ConnectionType, this.ProxyServer, "", 0)

If nInetHandle == 0
   this.ErrorNumber = GetLastError()
   Return("")
Endif



* Establecer una sesion
nHTTPSession = InternetConnect(nInetHandle, cServer, nPort, ;
                               cUsername, cPassword, ;
                               INTERNET_SERVICE_HTTP, 0, 0)

If nHTTPSession == 0
   this.ErrorNumber = GetLastError()
   InternetCloseHandle(nInetHandle)
   Return("")
Endif



* Conectarse
nHTTPResult = HttpOpenRequest(nHTTPSession, cVerb, ;
                              cURL, NULL, NULL, NULL, ;
                              INTERNET_FLAG_RELOAD + IIF(nPort == 443, INTERNET_FLAG_SECURE, 0), 0)

If nHTTPResult == 0
   this.ErrorNumber = GetLastError()
   InternetCloseHandle(nHTTPSession)
   InternetCloseHandle(nInetHandle)
   Return("")
Endif



* Establecer opciones de autenticacion con el proxy
If !this.SetProxyAuthentication(nHTTPResult)
   InternetCloseHandle(nHTTPResult)
   InternetCloseHandle(nHTTPSession)
   InternetCloseHandle(nInetHandle)
   Return("")
Endif



* Establecer opciones de certificados SSL
If this.IgnoreSSLErrors
   Declare Integer InternetSetOption In WinInet.DLL Integer hInternet, Integer dwFlags, Integer @dwValue, Integer cbSize

   nFlags = SECURITY_FLAG_IGNORE_UNKNOWN_CA + ;
            SECURITY_FLAG_IGNORE_CERT_DATE_INVALID + ;
            SECURITY_FLAG_IGNORE_CERT_CN_INVALID + ;
            SECURITY_FLAG_IGNORE_REVOCATION + ;
            SECURITY_FLAG_IGNORE_WRONG_USAGE

   If InternetSetOption(nHTTPResult, INTERNET_OPTION_SECURITY_FLAGS, @nFlags, 4) == 0
       this.ErrorNumber = GetLastError()
       InternetCloseHandle(nHTTPSession)
       InternetCloseHandle(nInetHandle)
       Return("")
    Endif
Endif



* Preparar la informacion a enviar
cPostBuffer = ""

If !this.Upload
   cHeaders = "Content-Type: application/x-www-form-urlencoded"

   For nCont = 1 to this.FieldCount
       cPostBuffer = cPostBuffer + this.PostFields[nCont, 2] + "=" + ;
                     this.URLEncode(this.PostFields[nCont, 3]) + "&"
   Next
else

   Do While .T.

      lFound = .F.
      cBoundary = this.TmpName(40)

      For nCont = 1 to this.FieldCount

          If At(cBoundary, this.URLEncode(this.PostFields[nCont, 2])) <> 0
             lFound = .T.
          Endif

          Do Case
             Case this.PostFields[nCont, 1] == 1
                   If At(cBoundary, this.URLEncode(this.PostFields[nCont, 3])) <> 0
                      lFound = .T.
                   Endif

             Case this.PostFields[nCont, 1] == 2
                   If At(cBoundary, FiletoStr(this.PostFields[nCont, 3])) <> 0
                      lFound = .T.
                   Endif

             Case this.PostFields[nCont, 1] == 3
                   If At(cBoundary, this.PostFields[nCont, 3]) <> 0
                      lFound = .T.
                   Endif

          EndCase

      Next

      If !lFound
         Exit
      Endif
   Enddo


   cHeaders = "Content-Type: multipart/form-data; boundary=" + cBoundary + Chr(13) + Chr(10)

   For nCont = 1 to this.FieldCount

       cPostBuffer = cPostBuffer + "--" + cBoundary + Chr(13) + Chr(10)
       cPostBuffer = cPostBuffer + 'Content-Disposition: form-data; name="' + this.PostFields[nCont, 2] + '";'

       Do Case
          Case this.PostFields[nCont, 1] == 1
               cPostBuffer = cPostBuffer + Chr(13) + Chr(10) + Chr(13) + Chr(10)
               cPostBuffer = cPostBuffer + this.URLEncode(this.PostFields[nCont, 3]) + Chr(13) + Chr(10)

          Case this.PostFields[nCont, 1] == 2
               cPostBuffer = cPostBuffer + ' filename="' + this.PostFields[nCont, 3] + '"' + Chr(13) + Chr(10)
               cPostBuffer = cPostBuffer + "Content-Type: application/upload" + Chr(13) + Chr(10) + Chr(13) + Chr(10)
               cPostBuffer = cPostBuffer + FiletoStr(this.PostFields[nCont, 3]) + Chr(13) + Chr(10)

          Case this.PostFields[nCont, 1] == 3
               cPostBuffer = cPostBuffer + ' filename="' + this.PostFields[nCont, 4] + '"' + Chr(13) + Chr(10)
               cPostBuffer = cPostBuffer + "Content-Type: application/upload" + Chr(13) + Chr(10) + Chr(13) + Chr(10)
               cPostBuffer = cPostBuffer + this.PostFields[nCont, 3] + Chr(13) + Chr(10)

       EndCase

   Next

   cPostBuffer = cPostBuffer + "--" + cBoundary  + "--" + Chr(13) + Chr(10)

Endif



* Enviar los datos
nRetval = HttpSendRequest(nHTTPResult, cHeaders, Len(cHeaders), ;
                          cPostBuffer, Len(cPostBuffer))

If nRetval = 0
   this.ErrorNumber = GetLastError()
   InternetCloseHandle(nHTTPSession)
   InternetCloseHandle(nInetHandle)
   Return("")
Endif



* Leer el resultado
cRes = ""

Do While .T.
   cReadBuffer = Space(65536)
   nBytesRead = 0

   nRetval = InternetReadFile(nHTTPResult, @cReadBuffer,;
                              Len(cReadBuffer), @nBytesRead)

   If nRetVal == 0
      Exit
   else
      If nBytesRead == 0
         Exit
      else
         cReadBuffer = IIF(nBytesRead > 1, SubStr(cReadBuffer, 1, nBytesRead), "")
         cRes = cRes + cReadBuffer
      Endif
   Endif

Enddo


cStatusCode = Space(10)
HttpQueryInfo(nHTTPResult, 19, @cStatusCode, Len(cStatusCode), 0)
this.StatusCode = Val(cStatusCode)


If nRetval = 0
   this.ErrorNumber = GetLastError()
   InternetCloseHandle(nHTTPSession)
   InternetCloseHandle(nInetHandle)
   Return("")
else
   InternetCloseHandle(nHTTPSession)
   InternetCloseHandle(nInetHandle)
Endif


Return(cRes)



***
* Function StartAsyncPost
***
Function StartAsyncPost(cServer, cURL, cUsername, cPassword, nPort, cVerb)
Local nPort, cPostBuffer, cHeaders, cBoundary, lFound


* Inicializar Variables
this.ErrorNumber = 0
this.StatusCode = 0

this.AsyncHTTPSession = 0
this.AsyncInetHandle = 0

If Type("cUserName") <> "C"
   cUserName = ""
Endif

If Type("cPassword") <> "C"
   cPassword = ""
Endif

If Type("nPort") <> "N"
   nPort = INTERNET_DEFAULT_HTTP_PORT
Endif

If Type("cVerb") <> "N"
   cVerb = "POST"
Endif



* Inicializar los servicios de Internet
this.AsyncInetHandle = InternetOpen("iFox", this.ConnectionType, this.ProxyServer, "", 0)

If this.AsyncInetHandle == 0
   this.ErrorNumber = GetLastError()
   Return(.F.)
Endif



* Establecer una sesion
this.AsyncHTTPSession = InternetConnect(this.AsyncInetHandle, cServer, nPort, ;
                                        cUsername, cPassword, ;
                                        INTERNET_SERVICE_HTTP, 0, 0)

If this.AsyncHTTPSession == 0
   this.ErrorNumber = GetLastError()
   InternetCloseHandle(this.AsyncInetHandle)
   Return(.F.)
Endif



* Conectarse
this.AsyncHTTPResult = HttpOpenRequest(this.AsyncHTTPSession, cVerb, ;
                                       cURL, NULL, NULL, NULL, ;
                                       INTERNET_FLAG_RELOAD + IIF(nPort == 443, INTERNET_FLAG_SECURE, 0), 0)

If this.AsyncHTTPResult == 0
   this.ErrorNumber = GetLastError()
   InternetCloseHandle(this.AsyncHTTPSession)
   InternetCloseHandle(this.AsyncInetHandle)
   Return(.F.)
Endif



* Establecer opciones de autenticacion con el proxy
If !this.SetProxyAuthentication(this.AsyncHTTPResult)
   InternetCloseHandle(this.AsyncHTTPResult)
   InternetCloseHandle(this.AsyncHTTPSession)
   InternetCloseHandle(this.AsyncInetHandle)
   Return(.F.)
Endif



* Preparar la informacion a enviar
cPostBuffer = ""

If !this.Upload
   cHeaders = "Content-Type: application/x-www-form-urlencoded"

   For nCont = 1 to this.FieldCount
       cPostBuffer = cPostBuffer + this.PostFields[nCont, 2] + "=" + ;
                     this.URLEncode(this.PostFields[nCont, 3]) + "&"
   Next
else

   Do While .T.

      lFound = .F.
      cBoundary = this.TmpName(40)

      For nCont = 1 to this.FieldCount

          If At(cBoundary, this.URLEncode(this.PostFields[nCont, 2])) <> 0
             lFound = .T.
          Endif

          Do Case
             Case this.PostFields[nCont, 1] == 1
                   If At(cBoundary, this.URLEncode(this.PostFields[nCont, 3])) <> 0
                      lFound = .T.
                   Endif

             Case this.PostFields[nCont, 1] == 2
                   If At(cBoundary, FiletoStr(this.PostFields[nCont, 3])) <> 0
                      lFound = .T.
                   Endif

             Case this.PostFields[nCont, 1] == 3
                   If At(cBoundary, this.PostFields[nCont, 3]) <> 0
                      lFound = .T.
                   Endif

          EndCase

      Next

      If !lFound
         Exit
      Endif
   Enddo


   cHeaders = "Content-Type: multipart/form-data; boundary=" + cBoundary + Chr(13) + Chr(10)

   For nCont = 1 to this.FieldCount

       cPostBuffer = cPostBuffer + "--" + cBoundary + Chr(13) + Chr(10)
       cPostBuffer = cPostBuffer + 'Content-Disposition: form-data; name="' + this.PostFields[nCont, 2] + '";'

       Do Case
          Case this.PostFields[nCont, 1] == 1
               cPostBuffer = cPostBuffer + Chr(13) + Chr(10) + Chr(13) + Chr(10)
               cPostBuffer = cPostBuffer + this.URLEncode(this.PostFields[nCont, 3]) + Chr(13) + Chr(10)

          Case this.PostFields[nCont, 1] == 2
               cPostBuffer = cPostBuffer + ' filename="' + this.PostFields[nCont, 3] + '"' + Chr(13) + Chr(10)
               cPostBuffer = cPostBuffer + "Content-Type: application/upload" + Chr(13) + Chr(10) + Chr(13) + Chr(10)
               cPostBuffer = cPostBuffer + FiletoStr(this.PostFields[nCont, 3]) + Chr(13) + Chr(10)

          Case this.PostFields[nCont, 1] == 3
               cPostBuffer = cPostBuffer + ' filename="' + this.PostFields[nCont, 4] + '"' + Chr(13) + Chr(10)
               cPostBuffer = cPostBuffer + "Content-Type: application/upload" + Chr(13) + Chr(10) + Chr(13) + Chr(10)
               cPostBuffer = cPostBuffer + this.PostFields[nCont, 3] + Chr(13) + Chr(10)

       EndCase

   Next

   cPostBuffer = cPostBuffer + "--" + cBoundary  + "--" + Chr(13) + Chr(10)

Endif


* Enviar los datos
If HttpAddRequestHeaders(this.AsyncHTTPResult, @cHeaders, Len(cHeaders), 536870912) == 0
   this.ErrorNumber = GetLastError()
   InternetCloseHandle(this.AsyncHTTPSession)
   InternetCloseHandle(this.AsyncInetHandle)
   Return(.F.)
Endif

cHeaders = "Content-Length: " + LTrim(Str(Len(cPostBuffer)))
If HttpAddRequestHeaders(this.AsyncHTTPResult, @cHeaders, Len(cHeaders), 536870912) == 0
   this.ErrorNumber = GetLastError()
   InternetCloseHandle(this.AsyncHTTPSession)
   InternetCloseHandle(this.AsyncInetHandle)
   Return(.F.)
Endif

If HttpSendRequestEx(this.AsyncHTTPResult, NULL, 0, 0, 0) == 0
   this.ErrorNumber = GetLastError()
   InternetCloseHandle(this.AsyncHTTPSession)
   InternetCloseHandle(this.AsyncInetHandle)
   Return(.F.)
Endif

this.AsyncData = cPostBuffer
this.TransferSize = Len(cPostBuffer)
this.CompletedTransfer = 0

this.EOT = .F.

Return(.T.)



***
* Function ContinueAsyncPost
***
Function ContinueAsyncPost()
Local cString, nRes

If Len(this.AsyncData) > 0
   cString = Left(this.AsyncData, this.BatchSize)
   this.AsyncData = Substr(this.AsyncData, this.BatchSize + 1)

   nRes = 0
   If InternetWriteFile(this.AsyncHTTPResult, @cString, Len(cString), @nRes) == 0
      Return(.F.)
   Endif

   If nRes == Len(cString)
      this.CompletedTransfer = this.CompletedTransfer + nRes
   else
      Return(.F.)
   Endif
Endif

If Len(this.AsyncData) == 0
   this.EOT = .T.
Endif   

Return(.T.)



***
* Function EndAsyncPost
***
Function EndAsyncPost()
Local cRes, cReadBuffer, nBytesRead, nRetVal, cStatusCode

this.AsyncData = ""

If HttpEndRequest(this.AsyncHTTPResult, 0, 1, 0) == 0
   this.ErrorNumber = GetLastError()
   InternetCloseHandle(this.AsyncHTTPSession)
   InternetCloseHandle(this.AsyncInetHandle)
   Return("")
Endif

cRes = ""
Do While .T.
   cReadBuffer = Space(65536)
   nBytesRead = 0

   nRetval = InternetReadFile(this.AsyncHTTPResult, @cReadBuffer,;
                              Len(cReadBuffer), @nBytesRead)

   If nRetVal == 0
      Exit
   else
      If nBytesRead == 0
         Exit
      else
         cReadBuffer = IIF(nBytesRead > 1, SubStr(cReadBuffer, 1, nBytesRead), "")
         cRes = cRes + cReadBuffer
      Endif
   Endif

Enddo


cStatusCode = Space(10)
HttpQueryInfo(this.AsyncHTTPResult, 19, @cStatusCode, Len(cStatusCode), 0)
this.StatusCode = Val(cStatusCode)


If nRetval = 0
   this.ErrorNumber = GetLastError()
   InternetCloseHandle(this.AsyncHTTPSession)
   InternetCloseHandle(this.AsyncInetHandle)
   Return("")
else
   InternetCloseHandle(this.AsyncHTTPSession)
   InternetCloseHandle(this.AsyncInetHandle)
Endif

Return(cRes)



***
* Function BuildPostData
***
Function BuildPostData()
Local cHeaders, cPostData
Local lFound, cBoundary, nCont

cPostData = ""

If !this.Upload
   cHeaders = "Content-Type: application/x-www-form-urlencoded"

   For nCont = 1 to this.FieldCount
       cPostData = cPostData + this.PostFields[nCont, 2] + "=" + ;
                     this.URLEncode(this.PostFields[nCont, 3]) + "&"
   Next
else

   Do While .T.

      lFound = .F.
      cBoundary = this.TmpName(40)

      For nCont = 1 to this.FieldCount

          If At(cBoundary, this.URLEncode(this.PostFields[nCont, 2])) <> 0
             lFound = .T.
          Endif

          Do Case
             Case this.PostFields[nCont, 1] == 1
                   If At(cBoundary, this.URLEncode(this.PostFields[nCont, 3])) <> 0
                      lFound = .T.
                   Endif

             Case this.PostFields[nCont, 1] == 2
                   If At(cBoundary, FiletoStr(this.PostFields[nCont, 3])) <> 0
                      lFound = .T.
                   Endif

             Case this.PostFields[nCont, 1] == 3
                   If At(cBoundary, this.PostFields[nCont, 3]) <> 0
                      lFound = .T.
                   Endif

          EndCase

      Next

      If !lFound
         Exit
      Endif
   Enddo


   cHeaders = "Content-Type: multipart/form-data; boundary=" + cBoundary + Chr(13) + Chr(10)

   For nCont = 1 to this.FieldCount

       cPostData = cPostData + "--" + cBoundary + Chr(13) + Chr(10)
       cPostData = cPostData + 'Content-Disposition: form-data; name="' + this.PostFields[nCont, 2] + '";'

       Do Case
          Case this.PostFields[nCont, 1] == 1
               cPostData = cPostData + Chr(13) + Chr(10) + Chr(13) + Chr(10)
               cPostData = cPostData + this.URLEncode(this.PostFields[nCont, 3]) + Chr(13) + Chr(10)

          Case this.PostFields[nCont, 1] == 2
               cPostData = cPostData + ' filename="' + this.PostFields[nCont, 3] + '"' + Chr(13) + Chr(10)
               cPostData = cPostData + "Content-Type: application/upload" + Chr(13) + Chr(10) + Chr(13) + Chr(10)
               cPostData = cPostData + FiletoStr(this.PostFields[nCont, 3]) + Chr(13) + Chr(10)

          Case this.PostFields[nCont, 1] == 3
               cPostData = cPostData + ' filename="' + this.PostFields[nCont, 4] + '"' + Chr(13) + Chr(10)
               cPostData = cPostData + "Content-Type: application/upload" + Chr(13) + Chr(10) + Chr(13) + Chr(10)
               cPostData = cPostData + this.PostFields[nCont, 3] + Chr(13) + Chr(10)

       EndCase

   Next

   cPostData = cPostData + "--" + cBoundary  + "--" + Chr(13) + Chr(10)

Endif

this.PostHeaders = cHeaders
this.PostData = cPostData

EndFunc



***
* Function URLEncode
***
Function URLEncode(cURL)
Local cResult, cChar, nSize, nCont
   
cResult = ""

For nCont = 1 to Len(cURL)
    cChar = SubStr(cURL, nCont, 1)

    If ATC(cChar, "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789") > 0
       cResult = cResult + cChar
       Loop
    Endif

    IF cChar == " "
       cResult = cResult + "+"
       Loop
    Endif
    
    cResult = cResult + "%" + Right(Transform(Asc(cChar), "@0"), 2)
Next

Return(cResult)



***
*   Function TmpName
*   Retorna un nombre unico 
***
Protected Function TmpName(nLen)
Local nCont, cNom_Arch

If Type("nLen") <> "N"
   nLen = 8
Endif

cNom_Arch = ""
For nCont = 1 TO nLen
    cNom_Arch = cNom_Arch + Chr(Irand(65, 90))
Next

Return(cNom_Arch)


EndDefine
