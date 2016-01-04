
If .F.
   ?  "Creado por Pablo Pioli"
Endif


#include IFOX.H


#define SOCKET_ERROR       -1 
#define INVALID_SOCKET     -1 

#define WSADATA_SIZE      398 
#define WS_VERSION        514
#define INADDR_NONE        -1

#define SOMAXCONN          0x7FFFFFFF


* Address Family Constants 
#define AF_INET             2
#define AF_IPX              6
#define AF_NETBIOS         17


* Socket Type Constants 
#define SOCK_STREAM         1 
#define SOCK_DGRAM          2 
#define SOCK_RAW            3 
#define SOCK_RDM            4 
#define SOCK_SEQPACKET      5 


* Protocol Constants 
#define IPPROTO_IP          0 
#define IPPROTO_TCP         6 
#define IPPROTO_UDP        17 
#define IPPROTO_RAW       255 



*************************   SOCKETS CLASS   ********************************

****************************************************************************
Define Class Sockets as iFox of iFox.PRG OlePublic

       StartOK = .F.
       SocketClosed = .F.

       Response = ""

       Hidden Sockets[1]
       Hidden UDPSockets[1, 3]
       Hidden ListenerSockets[1]

       PingAddress = ""
       PingTime = 0
       PingResponseSize = 0
       PingError = 0


***
* Function eInit
***
Protected Function eInit()
Local cData, lRes


* Declare Winsock 2 functions
this.StartOK = this.DeclareWinAPI()


* Start Sockets
If this.StartOK
   cData = Replicate(Chr(0), WSADATA_SIZE) 
   If WS_StartUp(WS_VERSION, @cData) <> 0
      this.StartOK = .F.
   Endif
Endif


EndFunc



***
* Function DeclareWinAPI
***
Hidden Function DeclareWinAPI()
Local lRes, cError
Private lStartSocketError


* Trap Errors (Users may have an older version of Winsock)
lStartSocketError = .F.

cError = On("Error")
On Error WS2_Error()


Declare Integer WSAStartup in WS2_32 as WS_StartUp Integer wVerRq, String @lpWSAData 
Declare Integer WSACleanup in WS2_32 as WS_CleanUp
Declare Integer WSAGetLastError in WS2_32 as WS_GetLastError

Declare Integer socket in WS2_32 as WS_CreateSocket Integer af, Integer Type, Integer Protocol 
Declare Integer closesocket in WS2_32 as WS_CloseSocket Integer hSocket

Declare Integer connect in WS2_32 as WS_Connect Integer hSocket, String sockaddr, Integer Len
Declare Integer send in WS2_32 as WS_Send Integer hSocket, String Buffer, Integer Len, Integer Flags
Declare Integer recv in WS2_32 as WS_Receive Integer hSocket, String Buffer, Integer Len, Integer Flags
Declare Integer select in WS2_32 as WS_Status Integer nfds, String @Read, String @Write, String @Errors, String TimeOut
Declare Integer sendto in WS2_32 as WS_SendTo Integer s, String @buf, Integer buflen, Integer wsflags, String @sendto, Integer tolen


Declare Long bind in WS2_32 as WS_Bind Integer hSocket, String SockAddr, Integer NameLen
Declare Long listen in WS2_32 as WS_Listen Integer hSocket, Long BackLog
Declare Long accept in WS2_32 as WS_Accept Integer hSocket, String @SockAddr, Long @NameLen

Declare Long getsockname in WS2_32 as WS_GetSocketName Integer hSocket, String @SockAddr, Long @NameLen
Declare Long getpeername in WS2_32 as WS_GetPeerName Integer hSocket, String @SockAddr, Long @NameLen

Declare Long inet_addr in WS2_32 as WS_IPtoNum String cp
Declare String inet_ntoa in WS2_32 as WS_NumtoIP Long nAddr

Declare Short htons in WS2_32 Long hostshort
Declare Long ntohs in WS2_32 Short x

Declare Integer gethostname in WS2_32 String @lpHostName, Integer iHostNameLenght
Declare Integer gethostbyname in WS2_32 String lpHostName
Declare Integer gethostbyaddr in WS2_32 Integer @addres, Integer len, Integer type

Declare Integer RtlMoveMemory in Win32API String @lpDest, Long nSource, Integer nBytes
Declare Integer RtlMoveMemory in Win32API As RtlMoveMemory2 String @lpDest, Long @nSource, Integer nBytes


On Error &cError


* Get Status
If lStartSocketError
   lRes = .F.
else
   lRes = .T.
Endif

Return(lRes)



***
* Function eDestroy
***
Protected Function eDestroy()

If !this.StartOK
   Return
Endif

WS_CleanUp()

EndFunc



***
* Function Connect
***
Function Connect(cHost, nPort)
Local nSocket, nAddress, cAddress

If !this.StartOK
   Return(0)
Endif

nSocket = 0
nAddress = this.GetAddressLong(cHost)

If nAddress <> INADDR_NONE

*!*	Declare Long WSASocket in WS2_32 Integer iAddressFamily, Integer iType, Integer iProtocol, String lpProtocolInfo, Long lpGroup, Long dwFlags
*!*	lcProtInfo = ""

*!*	x = ""
*!*	x = x + num2dword(0) &&|   DWORD            dwServiceFlags1;             0:4
*!*	x = x + num2dword(0) &&|   DWORD            dwServiceFlags2;             4:4
*!*	x = x + num2dword(0) &&|   DWORD            dwServiceFlags3;             8:4
*!*	x = x + num2dword(0) &&|   DWORD            dwServiceFlags4;            12:4
*!*	x = x + num2dword(0) &&|   DWORD            dwProviderFlags;            16:4
*!*	x = x + Replicate(" ", 16) &&|   GUID             ProviderId;                 20:16
*!*	x = x + num2dword(0) &&|   DWORD            dwCatalogEntryId;           36:4

*!*	for tt = 1 to 8
*!*	x = x + num2dword(0) &&|   WSAPROTOCOLCHAIN ProtocolChain;              40:32
*!*	Next

*!*	x = x + num2dword(0) &&|   int              iVersion;                   72:4
*!*	x = x + num2dword(AF_INET) &&|   int              iAddressFamily;             76:4  --
*!*	x = x + num2dword(0) &&|   int              iMaxSockAddr;               80:4
*!*	x = x + num2dword(0) &&|   int              iMinSockAddr;               84:4
*!*	x = x + num2dword(SOCK_STREAM) &&|   int              iSocketType;                88:4 --
*!*	x = x + num2dword(IPPROTO_TCP) &&|   int              iProtocol;                  92:4 ---
*!*	x = x + num2dword(0) &&|   int              iProtocolMaxOffset;         96:4
*!*	x = x + num2dword(0) &&|   int              iNetworkByteOrder;         100:4
*!*	x = x + num2dword(0) &&|   int              iSecurityScheme;           104:4 ---
*!*	x = x + num2dword(0) &&|   DWORD            dwMessageSize;             108:4
*!*	x = x + num2dword(0) &&|   DWORD            dwProviderReserved;        112:4
*!*	x = x + Replicate(" ", 256) &&|   TCHAR            szProt[WSAPROTOCOL_LEN+1]; 116:256


*!*	a = WSASocket(AF_INET, SOCK_STREAM, IPPROTO_TCP, x, 0, 0)

   nSocket = WS_CreateSocket(AF_INET, SOCK_STREAM, IPPROTO_TCP) 

   cAddress = ""
   cAddress = cAddress + APINumtoStr(AF_INET, 2)
   cAddress = cAddress + APINumtoStr(htons(nPort), 2)

   If nAddress >= 0
      cAddress = cAddress + APINumtoStr(nAddress, 4)
   else
      cAddress = cAddress + this.NumToLong(nAddress)
   Endif

   cAddress = cAddress + Replicate(Chr(0), 8)

   If WS_Connect(nSocket, cAddress, Len(cAddress)) == SOCKET_ERROR
      WS_CloseSocket(nSocket)
      nSocket = 0
   else
      Dimension this.Sockets[ALen(this.Sockets) + 1]
      this.Sockets[ALen(this.Sockets)] = nSocket
   Endif

Endif

Return(nSocket)



***
* Function Connect
***
Function ConnectUDP(cHost, nPort)
Local nSocket, nAddress, cAddress

If !this.StartOK
   Return(0)
Endif

nSocket = 0
nAddress = this.GetAddressLong(cHost)

If nAddress <> INADDR_NONE
   nSocket = WS_CreateSocket(AF_INET, SOCK_DGRAM, IPPROTO_UDP)

   Dimension this.UDPSockets[ALen(this.UDPSockets, 1) + 1, ALen(this.UDPSockets, 2)]
   this.UDPSockets[ALen(this.UDPSockets, 1), 1] = nSocket
   this.UDPSockets[ALen(this.UDPSockets, 1), 2] = cHost
   this.UDPSockets[ALen(this.UDPSockets, 1), 3] = nPort
Endif

Return(nSocket)



***
* Function GetAddressLong
***
Hidden Function GetAddressLong(cHost)
Local nRes, cIP

nRes = WS_IPtoNum(cHost)

If nRes == INADDR_NONE
   cIP = this.GetIPfromName(cHost)

   If At(",", cIP) <> 0
      cIP = Left(cIP, At(",", cIP) - 1)
   Endif

   nRes = WS_IPtoNum(cIP)
Endif

Return(nRes)



***
* Function Close
***
Function Close(nSocket)
Local nCont

If !this.StartOK
   Return
Endif

For nCont = 2 to ALen(this.Sockets)
    If this.Sockets[nCont] == nSocket
       ADel(this.Sockets, nCont)
       Dimension this.Sockets[ALen(this.Sockets) - 1]
       Exit
    Endif
Next

For nCont = 2 to ALen(this.ListenerSockets)
    If this.ListenerSockets[nCont] == nSocket
       ADel(this.ListenerSockets, nCont)
       Dimension this.ListenerSockets[ALen(this.ListenerSockets) - 1]
       Exit
    Endif
Next

WS_CloseSocket(nSocket)

EndFunc



***
* Function CloseUDP
***
Function CloseUDP(nSocket)
Local nCont

If !this.StartOK
   Return
Endif

For nCont = 2 to ALen(this.UDPSockets, 1)
    If this.UDPSockets[nCont, 1] == nSocket
       ADel(this.UDPSockets, nCont)
       Dimension this.UDPSockets[ALen(this.UDPSockets, 1) - 1, ALen(this.UDPSockets, 2)]
       Exit
    Endif
Next

WS_CloseSocket(nSocket)

EndFunc



***
* Function Listen
***
Function Listen(nPort, cIP)
Local nSocket, cAddress, lRes

If !this.StartOK
   Return(0)
Endif

If (Type("cIP") <> "C") .OR. (Empty(cIP))
   cIP = this.IPAddress()

   If At(",", cIP) <> 0
      cIP = Left(cIP, At(",", cIP) - 1)
   Endif
Endif

nSocket = WS_CreateSocket(AF_INET, SOCK_STREAM, IPPROTO_TCP) 

If nSocket > 0 Then

   lRes = .F.

   cAddress = ""
   cAddress = cAddress + APINumtoStr(AF_INET, 2)
   cAddress = cAddress + APINumtoStr(htons(nPort), 2)
   cAddress = cAddress + this.NumToLong(this.GetAddressLong(cIP))
   cAddress = cAddress + Replicate(APINumtoStr(0, 1), 8)

   If WS_Bind(nSocket, cAddress, Len(cAddress)) == 0
      If WS_Listen(nSocket, SOMAXCONN) == 0
      
         Dimension this.ListenerSockets[ALen(this.ListenerSockets) + 1]
         this.ListenerSockets[ALen(this.ListenerSockets)] = nSocket

         lRes = .T.
      Endif
   Endif

   If !lRes
      WS_CloseSocket(nSocket)
      nSocket = 0
   Endif

Endif

Return(nSocket)



***
* Function Send
***
Function Send(nSocket, cData)
Local lRes, cBuffer1, cBuffer2, cBuffer3, nCont


If !this.StartOK
   Return(.F.)
Endif


* Prepare response
lRes = .F.


* Build string buffer
cBuffer1 = APINumtoStr(1, 4)
cBuffer1 = cBuffer1 + APINumtoStr(nSocket, 4)

For nCont = 2 to 64
    cBuffer1 = cBuffer1 + APINumtoStr(0, 4)
Next


* Build auxiliar buffers
cBuffer2 = cBuffer1
cBuffer3 = cBuffer1


* Read data
nRes = WS_Status(0, @cBuffer1, @cBuffer2, @cBuffer3, APINumtoStr(2, 4) + APINumtoStr(0, 4))


* Process results
If nRes <> SOCKET_ERROR
   If APIStrtoNum(Left(cBuffer2, 4)) == 1
      If WS_Send(nSocket, cData, Len(cData), 0) == Len(cData)
         lRes = .T.
      Endif
   Endif
Endif

Return(lRes)



***
* Function SendUDP
***
Function SendUDP(nSocket, cData)
Local lFound, nCont, nAddress, cAddress


If !this.StartOK
   Return(.F.)
Endif


lFound = .F.
For nCont = 2 to ALen(this.UDPSockets, 1)
    If this.UDPSockets[nCont, 1] == nSocket
       lFound = .T.
       Exit
    Endif
Next

If !lFound
   Return(.F.)
Endif


nAddress = this.GetAddressLong(this.UDPSockets[nCont, 2])

cAddress = ""
cAddress = cAddress + APINumtoStr(AF_INET, 2)
cAddress = cAddress + APINumtoStr(htons(this.UDPSockets[nCont, 3]), 2)

If nAddress >= 0
   cAddress = cAddress + APINumtoStr(nAddress, 4)
else
   cAddress = cAddress + this.NumToLong(nAddress)
Endif

cAddress = cAddress + Replicate(Chr(0), 8)
   
Return(WS_SendTo(nSocket, @cData, Len(cData), 0, @cAddress, Len(cAddress)) <> -1)



***
* Function ReadUDP
***
Function ReadUDP(nSocket)
Return(this.Read(nSocket))



***
* Function Read
***
Function Read(nSocket)
Local cRes, cBuffer1, cBuffer2, cBuffer3, nCont
Local cReadBuffer, nRead


If !this.StartOK
   Return("")
Endif


* Prepare response
cRes = ""
this.SocketClosed = .F.


* Build string buffer
cBuffer1 = APINumtoStr(1, 4)
cBuffer1 = cBuffer1 + APINumtoStr(nSocket, 4)

For nCont = 2 to 64
    cBuffer1 = cBuffer1 + APINumtoStr(0, 4)
Next


* Build auxiliar buffers
cBuffer2 = cBuffer1
cBuffer3 = cBuffer1


* Read data
nRes = WS_Status(0, @cBuffer1, @cBuffer2, @cBuffer3, APINumtoStr(0, 4) + APINumtoStr(0, 4))


* Process results
If nRes <> SOCKET_ERROR

   * Readable sockets
   If APIStrtoNum(Left(cBuffer1, 4)) == 1

      cReadBuffer = Replicate(Chr(0), 8192)
      nRead = WS_Receive(nSocket, @cReadBuffer, Len(cReadBuffer), 0)

      If nRead > 0
         cRes = Left(cReadBuffer, nRead)
      else
         this.SocketClosed = .T.
      Endif

   Endif
Endif

Return(cRes)



***
* Function AcceptConnections
***
Function AcceptConnections(nSocket)
Local cBuffer1, cBuffer2, cBuffer3, nCont, nRes
Local cAux, nAux


If !this.StartOK
   Return(0)
Endif


* Build string buffer
cBuffer1 = APINumtoStr(1, 4)
cBuffer1 = cBuffer1 + APINumtoStr(nSocket, 4)

For nCont = 2 to 64
    cBuffer1 = cBuffer1 + APINumtoStr(0, 4)
Next


* Build auxiliar buffers
cBuffer2 = cBuffer1
cBuffer3 = cBuffer1


* Read data
nRes = WS_Status(0, @cBuffer1, @cBuffer2, @cBuffer3, APINumtoStr(0, 4) + APINumtoStr(0, 4))


* Process results
If nRes <> SOCKET_ERROR

   * Readable sockets
   If APIStrtoNum(Left(cBuffer1, 4)) == 1
      cAux = Replicate(Chr(0), 2 + 2 + 4 + 8)
      nAux = Len(cAux)
      nRes = WS_Accept(nSocket, @cAux, @nAux)
   else
      nRes = 0
   Endif

else
   nRes = 0
Endif

Return(nRes)



***
* Function SendReceive
***
Function SendReceive(nSocket, cData, nTimeOut, nDelayBetweenRetries)
Local lRes, tStart, cData

If !this.StartOK
   Return(.F.)
Endif

lRes = .F.
this.Response = ""

If Type("nTimeOut") <> "N"
   nTimeOut = 10
Endif

If Type("nDelayBetweenRetries") <> "N"
   nDelayBetweenRetries = 0
Endif

If this.Send(nSocket, cData)

   tStart = DateTime()
   Do While .T.

      cData = this.Read(nSocket)

      If !Empty(cData)
         this.Response = cData
         lRes = .T.
         Exit
      else
         If this.SocketClosed
            lRes = .F.
            Exit
         Endif
      Endif

      If nDelayBetweenRetries <> 0
         Win32Delay(nDelayBetweenRetries)
      Endif

      If DateTime() - tStart > nTimeOut
         lRes = .F.
         Exit
      Endif

   Enddo

Endif

Return(lRes)



***
* Function WaitFor
***
Function WaitFor(nSocket, cPattern, nTimeOut)
Local lRes, tStart, cData

If !this.StartOK
   Return(.F.)
Endif

lRes = .F.


If Type("nTimeOut") <> "N"
   nTimeOut = 10
Endif

If Type("cPattern") <> "C"
   cPattern = ""
Endif


this.Response = ""
tStart = DateTime()

Do While .T.

   cData = this.Read(nSocket)

   If !Empty(cData)
      this.Response = this.Response + cData

      If (Len(cPattern) == 0) .OR. (AT(cPattern, this.Response) <> 0)
         lRes = .T.
         Exit
      Endif
   else
      If this.SocketClosed
         lRes = .F.
         Exit
      Endif
   Endif

   If DateTime() - tStart > nTimeOut
      lRes = .F.
      Exit
   Endif

Enddo

Return(lRes)



***
* Function WaitForSize
***
Function WaitForSize(nSocket, nSize, nTimeOut)
Local lRes, tStart, cData

If !this.StartOK
   Return(.F.)
Endif

lRes = .F.

If Type("nTimeOut") <> "N"
   nTimeOut = 10
Endif

this.Response = ""
tStart = DateTime()

Do While .T.

   cData = this.Read(nSocket)

   If !Empty(cData)
      this.Response = this.Response + cData

      If Len(this.Response) >= nSize
         lRes = .T.
         Exit
      Endif
   else
      If this.SocketClosed
         lRes = .F.
         Exit
      Endif
   Endif

   If DateTime() - tStart > nTimeOut
      lRes = .F.
      Exit
   Endif

Enddo

Return(lRes)



***
* Function GetLocalHost
***
Function GetLocalHost(nSocket)
Local nCont, cRes, cAddress, nLen

If !this.StartOK
   Return("")
Endif

For nCont = 2 to ALen(this.ListenerSockets)
    If this.ListenerSockets[nCont] == nSocket
       Return("")
    Endif
Next

cRes = ""

cAddress = Replicate(Chr(0), 2 + 2 + 4 + 8)
nLen = Len(cAddress)

If WS_GetSocketName(nSocket, @cAddress, @nLen) <> SOCKET_ERROR
   cRes = WS_NumtoIP(APIStrtoNum(SubStr(cAddress, 5, 4)))
Endif

Return(cRes)



***
* Function GetLocalHostIP
***
Function GetLocalHostIP(nSocket)
Local nCont, cRes, cAddress, nLen

If !this.StartOK
   Return("")
Endif

For nCont = 2 to ALen(this.ListenerSockets)
    If this.ListenerSockets[nCont] == nSocket
       Return("")
    Endif
Next

cRes = ""

cAddress = Replicate(Chr(0), 2 + 2 + 4 + 8)
nLen = Len(cAddress)

If WS_GetSocketName(nSocket, @cAddress, @nLen) <> SOCKET_ERROR
   cRes = WS_NumtoIP(APIStrtoNum(SubStr(cAddress, 5, 4)))
Endif

Return(cRes)



***
* Function GetRemoteHost
***
Function GetRemoteHost(nSocket)
Local nCont, cRes, cAddress, nLen

If !this.StartOK
   Return("")
Endif

For nCont = 2 to ALen(this.ListenerSockets)
    If this.ListenerSockets[nCont] == nSocket
       Return("")
    Endif
Next

For nCont = 2 to ALen(this.ListenerSockets)
    If this.ListenerSockets[nCont] == nSocket
       Return("")
    Endif
Next

cRes = ""

cAddress = Replicate(Chr(0), 2 + 2 + 4 + 8)
nLen = Len(cAddress)

If WS_GetPeerName(nSocket, @cAddress, @nLen) <> SOCKET_ERROR
   cRes = WS_NumtoIP(APIStrtoNum(SubStr(cAddress, 5, 4)))
Endif

Return(cRes)



***
* Function GetRemoteHostIP
***
Function GetRemoteHostIP(nSocket)
Local nCont, cRes, cAddress, nLen

If !this.StartOK
   Return("")
Endif

For nCont = 2 to ALen(this.ListenerSockets)
    If this.ListenerSockets[nCont] == nSocket
       Return("")
    Endif
Next

cRes = ""

cAddress = Replicate(Chr(0), 2 + 2 + 4 + 8)
nLen = Len(cAddress)

If WS_GetPeerName(nSocket, @cAddress, @nLen) <> SOCKET_ERROR
   cRes = WS_NumtoIP(APIStrtoNum(SubStr(cAddress, 5, 4)))
Endif

Return(cRes)



***
* Function GetPort
***
Function GetPort(nSocket)
Local nCont, nRes, cAddress, nLen

If !this.StartOK
   Return(-1)
Endif

cAddress = Replicate(Chr(0), 2 + 2 + 4 + 8)
nLen = Len(cAddress)

If WS_GetSocketName(nSocket, @cAddress, @nLen) <> SOCKET_ERROR
   nRes = ntohs(APIStrtoNum(SubStr(cAddress, 3, 4)))
else
   nRes = -1
Endif

Return(nRes)



***
* Function IPAddress
***
Function IPAddress()
Local cRes

If !this.StartOK
   Return("")
Endif

cHostName = Space(256)
If gethostname(@cHostName, 256) <> SOCKET_ERROR
   cRes = this.GetIPFromName(cHostName)
else
   cRes = ""
Endif

Return(cRes)



***
* Function GetIPFromName
***
Function GetIPFromName(cHost)
Local cRes, nHostAddr, cWSHostEnt
Local nHostEnt_AddrList, nHostEnt_Lenght, cHostIP_Addr, nHostIP_Addr
Local cTempIP_Addr, nCont

If !this.StartOK
   Return("")
Endif

cRes = ""
nHostAddr = gethostbyname(AllTrim(cHost))

If nHostAddr <> 0

   cWSHostEnt = Replicate(Chr(0), 4 + 4 + 2 + 2 + 4)
   RtlMoveMemory(@cWSHostEnt, nHostAddr, 16)
         
   nHostEnt_AddrList = this.StrToLong(SubStr(cWSHostEnt, 13, 4))
   nHostEnt_Lenght = this.StrToInt(SubStr(cWSHostEnt, 11, 2))


   Do While .T.

      cHostIP_Addr = Replicate(Chr(0), 4)
      RtlMoveMemory(@cHostIP_Addr, nHostEnt_AddrList, 4)

      nHostIP_Addr = this.StrToLong(cHostIP_Addr)

      If nHostIP_Addr == 0
         Exit
      else
         cRes = cRes + IIF(Empty(cRes), "", ",")
      Endif

      cTempIP_Addr = Replicate(Chr(0), nHostEnt_Lenght)

      RtlMoveMemory(@cTempIP_Addr, nHostIP_Addr, nHostEnt_Lenght)

      For nCont = 1 to nHostEnt_Lenght
          cRes = cRes + Transform(Asc(SubStr(cTempIP_Addr, nCont, 1))) + ;
                        IIF(nCont = nHostEnt_Lenght, "", ".")
      Next

      nHostEnt_AddrList = nHostEnt_AddrList + 4
   Enddo
Endif

Return(cRes)



***
* Function GetNameFromIP
***
Function GetNameFromIP(cIP)
Local nIP, nHostEnt, cHostEnt

If !this.StartOK
   Return("")
Endif

cRes = ""

nIP = WS_IPtoNum(cIP) 
nHostEnt = gethostbyaddr(@nIP, 4, AF_INET) 

If nHostEnt <> 0
   cHostEnt = this.GetMemBuf(nHostEnt, 16) 
   cRes = this.GetMemStr(this.Buf2DWord(SubStr(cHostEnt, 1, 4)))
Endif

Return(cRes)



***
* Function StrToLong
***
Hidden Function StrToLong(cLongStr)
Local nCont, nRes, cLongStr

nRes = 0
cLongStr = IIF(Empty(cLongStr), "", cLongStr)

For nCont = 0 to 24 Step 8
    nRes  = nRes + (Asc(cLongStr) * (2 ^ nCont))
    cLongStr = Right(cLongStr, Len(cLongStr) - 1)
Next

Return(nRes)



***
* Function StrToInt
***
Hidden Function StrToInt(cIntStr)
Local nCont, nRes, cIntStr

nRes = 0
cIntStr = IIF(Empty(cIntStr), "", cIntStr)

For nCont = 0 to 8 Step 8
    nRes = nRes + (Asc(cIntStr) * (2 ^ nCont))
    cIntStr = Right(cIntStr, Len(cIntStr) - 1)
Next

Return(nRes)



***
* Function GetMemBuf
***
Hidden Function GetMemBuf(nAddr, nSize) 
Local cBuffer

cBuffer = Replicate(Chr(0), nSize)
RtlMoveMemory(@cBuffer, @nAddr, nSize) 

Return(cBuffer)



***
* Function GetMemStr
***
Hidden Function GetMemStr(nAddr) 
Local cBuffer

cBuffer = this.GetMemBuf(nAddr, 250) 

Return(SubStr(cBuffer, 1, At(Chr(0), cBuffer) - 1 ))



***
* Function Buf2DWord
***
Hidden Function Buf2DWord(cBuffer) 
Local cRes

cRes = Asc(SubStr(cBuffer, 1, 1)) + ; 
       Asc(SubStr(cBuffer, 2, 1)) * 256 + ;
       Asc(SubStr(cBuffer, 3, 1)) * 65536 + ;
       Asc(SubStr(cBuffer, 4, 1)) * 16777216 

Return(cRes)



***
* Function NumToLong
*   Similar a APINumtoStr pero maneja numeros negativos
***
Hidden Function NumToLong(nNumber)
Local cRes

cRes = Space(4)
RtlMoveMemory2(@cRes, @nNumber, 4)

Return(cRes)



***
* Function Ping
***
Function Ping(cHost, nDataSize, nTimeOut)
Local lRes, nAddress, nPort, cData, cBuffer, nBufferLen, cDataRes


* Inicializar propiedades
If (Type("nDataSize") <> "N") .OR. (nDataSize <= 0)
   nDataSize = 32
Endif

If (Type("nTimeOut") <> "N") .OR. (nTimeOut == 0)
   nTimeOut = 500
Endif

this.PingAddress = ""
this.PingTime = 0
this.PingResponseSize = 0
this.PingError = 0


* Declare ICMP Functions
Declare Long IcmpCreateFile in "icmp.dll"

Declare Long IcmpCloseHandle in "icmp.dll" Long IcmpHandle
   
Declare Long IcmpSendEcho in "icmp.dll" ;
        Long IcmpHandle, ;
        Long DestinationAddress, ;
        String RequestData, ;
        Long RequestSize, ;
        Long RequestOptions, ;
        String ReplyBuffer, ;
        Long ReplySize, ;
        Long Timeout


* Ping
nAddress = this.GetAddressLong(cHost)

If nAddress <> INADDR_NONE
   nPort = IcmpCreateFile()

   If nPort <> 0

      cBuffer = ""
      cBuffer = cBuffer + APINumtoStr(0, 4)
      cBuffer = cBuffer + APINumtoStr(0, 4)
      cBuffer = cBuffer + APINumtoStr(0, 4)

      cBuffer = cBuffer + APINumtoStr(0, 2)
      cBuffer = cBuffer + APINumtoStr(0, 2)

      cBuffer = cBuffer + APINumtoStr(0, 4)

      cBuffer = cBuffer + APINumtoStr(0, 1)
      cBuffer = cBuffer + APINumtoStr(0, 1)
      cBuffer = cBuffer + APINumtoStr(0, 1)
      cBuffer = cBuffer + APINumtoStr(0, 1)
      cBuffer = cBuffer + APINumtoStr(0, 4)

      cBuffer = cBuffer + Replicate(Chr(0), 250)
      nBufferLen = Len(cBuffer)

      cData = Left(Replicate("iFox Ping", 100), nDataSize)

      If IcmpSendEcho(nPort, nAddress, cData, Len(cData), 0, ;
                      @cBuffer, @nBufferLen, nTimeOut) <> 0

         this.PingAddress = WS_NumtoIP(APIStrToNum(Left(cBuffer, 4)))
         this.PingTime = APIStrToNum(SubStr(cBuffer, 9, 4))
         this.PingResponseSize = APIStrToNum(SubStr(cBuffer, 13, 2))

         If this.PingResponseSize == nDataSize
            cDataRes = SubStr(cBuffer, 29, 250)
            If At(Chr(0), cDataRes) <> 0
               cDataRes = Left(cDataRes, At(Chr(0), cDataRes) - 1)
            Endif

            If cDataRes == cData
               lRes = .T.
            else
               this.PingError = -102
            Endif
         else
            this.PingError = -101
         Endif
      else
         this.PingError = APIStrToNum(SubStr(cBuffer, 5, 4))
      Endif

      IcmpCloseHandle(nPort)
   Endif

Endif


Return(lRes)


EndDefine




*************************   AUXILIAR FUNCTIONS   ***************************

****************************************************************************

***
* Function WS2_Error
*   Error handling function
***
Function WS2_Error()
lStartSocketError = .T.
EndFunc
