
If .F.
   ?  "Creado por Pablo Pioli"
Endif


#include IFOX.H


**************************   IFOX CLASS   **********************************

****************************************************************************
Define Class iFox as Custom


***
* Function Init
*   Initialize iFox
***
Hidden Function Init()


* Disable wait states
Sys(2335, 0)


* Initialize environment

* These a the default values for VFP7
* kept for VFP6 compatibility
Set Exclusive Off
Set Talk Off
Set Safety Off

* General settings
Set Deleted On
Set Date to DMY
Set Multilocks On
Set Century to 19 RollOver 80
Set Century Off
Set Near On
Set Decimals to 2
Set Separator to "."
Set Point to ","
Set UDFParms to Value
Set ANSI On
Set Exact On
Set Escape Off
Set Reprocess to 1
Set Refresh to 5, 30
Set Collate to "Spanish"
Set Hours to 24
Set Notify Off
Set CPDialog Off
Set LogErrors Off


* Initialize properties


* Set a reference to the axuiliar files
Set Procedure to AUXILIAR.PRG Additive


* This way any object can call another without the overhead
* of going thru COM
Set Procedure to SOCKETS.PRG Additive
Set Procedure to FTP.PRG Additive
Set Procedure to HTTP.PRG Additive
Set Procedure to POP3.PRG Additive
Set Procedure to SMTP.PRG Additive
Set Procedure to CONNECT.PRG Additive


* Initialize random number generator
Rand(-1)


* So each subclass can make it's own initialization
this.eInit()


EndFunc



***
* Function eInit
***
Protected Function eInit()
EndFunc



***
* Function Destroy
***
Hidden Function Destroy()

this.eDestroy()

EndFunc



***
* Function eDestroy
***
Protected Function eDestroy()
EndFunc



***
* Function GetVersion()
***
Function GetVersion()
Local cRes, cName, aVersion

cName = SubStr(Sys(16, 0), 23)

Dimension aVersion[1]
If AGetFileVersion(aVersion, cName) >= 4
   cRes = aVersion[4]
else
   cRes = "Error"
Endif

Return(cRes)





#ifdef FINAL_VERSION
***
* Function Error
***
Protected Function Error(nError, cMethod, nLine, cDesc)
Local cMssg, cCode, cErrorFile


* Get error message description
If Type("cDesc") <> "C"
   cDesc = Message()
Endif


* Generate error text
cMssg = "Número de error: " + LTrim(Str(nError, 6, 0)) + " - "
cMssg = cMssg + "Método: " + cMethod + " - "
cMssg = cMssg + "Línea: " + LTrim(Str(nLine, 6, 0)) + " - "
cMssg = cMssg + "Descripción: " + cDesc


* Generate a COM error and stop execution
ComReturnError("iFox", cMssg)

EndFunc
#endif


EndDefine
