
If .F.
   ?  "Creado por Pablo Pioli"
Endif


#include IFOX.H


***************************   CONNECT CLASS   ******************************

****************************************************************************
Define Class Connect as iFox of iFox.PRG OlePublic

       ForceConnection = .F.
       Dimension Connections[1]


***
* Function eInit
***
Protected Function eInit()


* Registrar APIs
Declare Integer InternetDial in WinInet ;
        Integer nHandle, ;
        String @lpcDialUp, ;
        Integer nAccessType, ;
        Integer @nHandle, ;
        Integer nFlags

Declare Integer InternetHangUp in WinInet ;
        Integer nHandle, ;
        Integer nFlags

Declare Integer InternetAutodial in WinInet ;
        Integer nAccessType, ;
        Integer nHandle

Declare Integer InternetAutodialHangup in WinInet ;
        Integer nFlags

Declare Integer InternetGetConnectedState in WinInet ;
         Integer @nFlags, ;
         Integer nReserved

Declare Long RasEnumEntries in "rasapi32.dll" ;
        String Reserved, ;
        String Phonebook, ;
        String @RASEntryName, ;
        Long @Len, ;
        Long @Entries

Declare Long RasEnumConnections in "rasapi32.dll" ;
        String @Connections, ;
        Long @Size, ;
        Long @Result


EndFunc



***
* Function EnumEntries
***
Function EnumEntries()
Local cRes, cEntries, nEntries, nLen, nCont, cString

cRes = ""

cEntries = APINumtoStr(264, 4) + Replicate(Chr(0), 260)

For nCont = 1 to 254
    cEntries = cEntries + APINumtoStr(0, 4) + Replicate(Chr(0), 260)
Next

nLen = Len(cEntries)
nEntries = 0


If RasEnumEntries(NULL, NULL, @cEntries, @nLen, @nEntries) == 0
   For nCont = 1 to nEntries
       cString = SubStr(cEntries, (nCont - 1) * 264 + 1 + 4, 256)
       If At(Chr(0), cString) <> 0
          cString = Left(cString, At(Chr(0), cString) - 1)
       Endif

       If Len(cRes) <> 0
          cRes = cRes + Chr(13) + Chr(10)
       Endif

       cRes = cRes + cString
   Next
Endif

Return(cRes)



***
* Function EnumConnections
***
Function EnumConnections()
Local cConnections, nConnections, nLen, nCont, cString

cConnections = APINumtoStr(412, 4) + Replicate(Chr(0), 408)

For nCont = 1 to 254
    cConnections = cConnections + APINumtoStr(0, 4) + Replicate(Chr(0), 408)
Next

nLen = Len(cConnections)
nConnections = 0


If RasEnumConnections(@cConnections, @nLen, @nConnections) == 0
   For nCont = 1 to nConnections

       Dimension this.Connections[nCont]
       this.Connections[nCont] = CreateObject("Connection")

       cString = SubStr(cConnections, (nCont - 1) * 412 + 1 + 8, 256)
       If At(Chr(0), cString) <> 0
          cString = Left(cString, At(Chr(0), cString) - 1)
       Endif

       this.Connections[nCont].EntryName = cString

       cString = SubStr(cConnections, (nCont - 1) * 412 + 1 + 4, 4)
       this.Connections[nCont].Handle = APIStrtoNum(cString)

   Next
else
   nConnections = 0
Endif

Return(nConnections)



***
* Function Dial
***
Function Dial(cConnection)
Local nRes, nFlags

If (Type("cConnection") <> "C") .OR. (Empty(cConnection))
   cConnection = ""
Endif

If this.ForceConnection
   nFlags = 2
else
   nFlags = 0
Endif

If Empty(cConnection)
   If InternetAutodial(nFlags, 0) <> 0
      nRes = -1
   else
      nRes = 0
   Endif
else
   nRes = 0
   If InternetDial(0, @cConnection, nFlags, @nRes, 0) <> 0
      nRes = 0
   Endif
Endif

Return(nRes)



***
* Function HangUp
***
Function HangUp(nConnection)
Local lRes

If Type("nConnection") <> "N"
   nConnection = 0
Endif

If Empty(nConnection)
   If InternetAutodialHangup(0) == 0
      lRes = .F.
   else
      lRes = .T.
   Endif
else
   If InternetHangUp(@nConnection, 0) == 0
      lRes = .T.
   else
      lRes = .F.
   Endif
Endif

Return(lRes)



***
* Function IsConnected
***
Function IsConnected()
Local lRes, nState

nState = 0
If InternetGetConnectedState(@nState, 0) <> 0
   lRes = .T.
else
   lRes = .F.
Endif

Return(lRes)



***
* Function Statistics
***
Function Statistics(nConnection, nType)
Local nRes, nVersion, oUtils
Local cBuffer, nCont

nRes = 0
nVersion = Val(OS(3) + "." + OS(4))

Do Case
   Case nVersion >= 5
        Declare Integer RasGetConnectionStatistics in "rasapi32.dll" ;
                Long Connection, ;
                String @Statistics

        cBuffer = ""
        For nCont = 1 to 14
            cBuffer = cBuffer + APINumtoStr(0, 4)
        Next

        cBuffer = APINumtoStr(Len(cBuffer) + 4, 4) + cBuffer

        If RasGetConnectionStatistics(nConnection, @cBuffer) == 0

           Do Case
              Case nType == 1
                   nRes = APIStrtoNum(SubStr(cBuffer, 13 * 4 + 1, 4))

              Case nType == 2
                   nRes = APIStrtoNum(SubStr(cBuffer,  2 * 4 + 1, 4))

              Case nType == 3
                   nRes = APIStrtoNum(SubStr(cBuffer,  1 * 4 + 1, 4))

              Case nType == 4
                   nRes = APIStrtoNum(SubStr(cBuffer, 14 * 4 + 1, 4))

              OtherWise
                   nRes = 0

           EndCase

        else
           nRes = 0
        Endif


   Case At("NT", Upper(OS())) <> 0
        nRes = 0


   OtherWise
        oUtils = CreateObject("Utils")

        Do Case
           Case nType == 1
                If oUtils.GetRegValue(HKEY_DYN_DATA, "PerfStats\StatData", "Dial-Up Adapter\ConnectSpeed")
                   nRes = APIStrtoNum(oUtils.Response)
                else
                   nRes = 0
                Endif


           Case nType == 2
                If oUtils.GetRegValue(HKEY_DYN_DATA, "PerfStats\StatData", "Dial-Up Adapter\BytesRecvd")
                   nRes = APIStrtoNum(oUtils.Response)
                else
                   nRes = 0
                Endif


           Case nType == 3
                If oUtils.GetRegValue(HKEY_DYN_DATA, "PerfStats\StatData", "Dial-Up Adapter\BytesXmit")
                   nRes = APIStrtoNum(oUtils.Response)
                else
                   nRes = 0
                Endif


           Case nType == 4
                nRes = 0


           OtherWise
                nRes = 0

        EndCase
EndCase

Return(nRes)


EndDefine



************************   CONNECTION CLASS   ******************************

****************************************************************************
Define Class Connection as Custom

             EntryName = ""
             Handle = 0

EndDefine
