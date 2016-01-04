
If .F.
   ?  "Creado por Pablo Pioli"
Endif


#include IFOX.H



***************************   NTP CLASS   **********************************

****************************************************************************
Define Class NTP as iFox of iFox.PRG OlePublic

       Hidden Sockets
       Server = ""
       Port = 123
       ErrorNumber = 0

       Date = CtoD("")
       Time = ""
       AdjustedDate = CtoD("")
       AdjustedTime = ""


***
* Function eInit()
***
Function eInit()

this.Sockets = CreateObject("Sockets")

EndFunc


       

***
* Function GetTime()
***
Function GetTime()
Local nSocket, cData, tStartTime, cRes, tTime
Local oUtils, nCurrentTimeZone


* Crear el socket
nSocket = this.Sockets.ConnectUDP(this.Server, this.Port)

this.ErrorNumber = 0

If nSocket < 0
   this.ErrorNumber = 1
   Return(.F.)
Endif
 
 
* Enviar la solicitud
cData = Chr(27) + Replicate("0", 47)
If !this.Sockets.SendUDP(nSocket, cData)
   this.ErrorNumber = 2
   this.Sockets.CloseUDP(nSocket)
   Return(.F.)
Endif


* Esperar la respuesta
tStartTime = DateTime()
cRes = ""
Do While DateTime() - tStartTime < 5
   cData = this.Sockets.Read(nSocket)
   If Len(cData) <> 0
      cRes = cRes + cData
      If Len(cRes) >= 48
         Exit
      Endif
   Endif
Enddo

this.Sockets.CloseUDP(nSocket)


* Procesar la respuesta
If Len(cRes) < 48
   this.ErrorNumber = 3
   Return(.F.)
Endif

tTime = DateTime(1900, 1, 1)
tTime = tTime + this.GetMilliSeconds(cRes) / 1000

this.Date = Date(Year(tTime), Month(tTime), Day(tTime))
this.Time = PadL(Int(Hour(tTime)), 2, "0") + ":" + ;
            PadL(Int(Minute(tTime)), 2, "0") + ":" + ;
            PadL(Int(Sec(tTime)), 2, "0")


* Ajustar el time zone
this.AdjustedDate = this.Date

oUtils = CreateObject("Utils")
nCurrentTimeZone = oUtils.GetTimeZoneOffsetMinutes()

If nCurrentTimeZone <> 0
   tTime = tTime + nCurrentTimeZone * 60

   this.AdjustedDate = Date(Year(tTime), Month(tTime), Day(tTime))
   this.AdjustedTime = PadL(Int(Hour(tTime)), 2, "0") + ":" + ;
                       PadL(Int(Minute(tTime)), 2, "0") + ":" + ;
                       PadL(Int(Sec(tTime)), 2, "0")
else
   this.AdjustedTime = this.Time
Endif

Return(.T.)




***
* Hidden Function GetMilliSeconds
***
Hidden Function GetMilliSeconds(cData)
Local nIntPart, nFractPart, nCont

cData = SubStr(cData, 41)
If Len(cData) > 8
   cData = Left(cData, 8)
Endif

nIntPart = 0
nFractPart = 0

For nCont = 1 to 4
    nIntPart = 256 * nIntPart + Asc(SubStr(cData, nCont, 1))
Next

For nCont = 5 to 8
    nFractPart = 256 * nFractPart + Asc(SubStr(cData, nCont, 1))
Next

Return(nIntPart * 1000 + (nFractPart * 1000) / 0x100000000)


EndDefine
