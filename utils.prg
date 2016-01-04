
If .F.
   ?  "Creado por Pablo Pioli"
Endif


#include IFOX.H


***************************   UTILS CLASS   *******************************

****************************************************************************
Define Class Utils as Custom OlePublic

       Response = ""


***
* Function Init
***
Protected Function Init()

* Declare WinAPI functions
Declare Integer CoCreateGuid In OLE32.DLL String @cGUID

Declare Integer GetTimeZoneInformation in Win32API ; 
        String @TimeZoneInformation 

Declare RtlMoveMemory in WIN32API as APIMoveMemory ;
        Integer @DestNum, ;
        String @pVoidSource, ;
        Integer nLength

Declare Integer RegOpenKeyEx in Win32API ;
        Integer nHKey, String cSubKey, Long Options, ;
        Integer SecurityMask, Long @nResult

Declare Integer RegCloseKey in Win32API ;
        Integer nHKey

Declare Integer RegQueryValueEx in Win32API ;
        Integer nHKey, String lpszValueName, Integer dwReserved, ;
        Integer @lpdwType, String @lpbData, Integer @lpcbData

Declare Integer GetPrivateProfileString in Win32API ;
        as GetPrivateINI String, String, String, String, integer, String

Declare Integer GetProfileString in Win32API ;
        as GetWinINI String, String, String, String, integer

EndFunc



***
*  Function GenGUID
***
Function GenGUID()
Local cGUID, cRes, nCont
Local cData1, cData2, cData3, cData4, cData5

cGUID = Replicate(Chr(0), 17)
cRes = ""

If CoCreateGuid(@cGUID) == 0

   cData1 = Right(Transform(this.StrToLong(Left(cGUID, 4)), "@0"), 8)
   cData2 = Right(Transform(this.StrToLong(SubStr(cGUID, 5, 2)), "@0"), 4)
   cData3 = Right(Transform(this.StrToLong(SubStr(cGUID, 7, 2)), "@0"), 4)
   cData4 = Right(Transform(this.StrToLong(SubStr(cGUID, 9, 1)), "@0"), 2) + ;
            Right(Transform(this.StrToLong(SubStr(cGUID, 10, 1)), "@0"), 2)

   cData5 = ""
   For nCont = 1 TO 6
       cData5 = cData5 + Right(Transform(this.StrToLong(SubStr(cGUID, 10 + nCont, 1))), 2)
   Next

   IF Len(cData5) < 12
      cData5 = cData5 + Replicate("0", 12 - Len(cData5))
   Endif

   cRes = "{" + cData1 + "-" + cData2 + "-" + cData3 + "-" + cData4 + "-" + cData5 + "}"
Endif

Return(cRes)



***
* Function StrToLong
***
Hidden Function StrToLong(cLong)
Local nRes, nCont

nRes = 0
For nCont = 0 to 24 Step 8
   nRes = nRes + (Asc(cLong) * (2 ^ nCont))
   cLong = Right(cLong, Len(cLong) - 1)
Next

Return(nRes)



***
* Function NeedEncoding
***
Function NeedEncoding(cString)
Local lRes, nCont

lRes = .F.

For nCont = 1 to Len(cString)
    If Asc(SubStr(cString, nCont, 1)) > 127
       lRes = .T.
       Exit
    Endif
Next

Return(lRes)



***
* Function Encode
***
Function Encode(cString, nMode, nExtraChar)
Local cRes, nCont, cLetter, nPos, cAux, cLine

cRes = ""

Do Case
   Case nMode == 1  && Codifica los encabezados en Quoted Printable
        nPos = 0 

        For nCont = 1 to Len(cString)

            cLetter = SubStr(cString, nCont, 1)

            Do Case
               Case (Asc(cLetter) >= 33) .AND. (Asc(cLetter) <= 60)
                    * No transformation needed

               Case (Asc(cLetter) >= 62) .AND. (Asc(cLetter) <= 126)
                    * No transformation needed

               Case (Asc(cLetter) == 13) .AND. (Asc(SubStr(cString, nCont + 1, 1)) == 10)
                    cLetter = cLetter + Chr(10)
                    nCont = nCont + 1

               Case (Asc(cLetter) == 32) .OR. (Asc(cLetter) == 9)
                    * No transformation needed (for now)

               OtherWise
                    cLetter = "=" + Right(Transform(Asc(cLetter), "@0"), 2)

            EndCase


            If cLetter = Chr(13) + Chr(10)
               If (Asc(Right(cRes, 1)) == 32) .OR. (Asc(Right(cRes, 1)) == 9)
                  cAux = "=" + Right(Transform(Asc(Right(cRes, 1)), "@0"), 2)
                  cRes = Left(cRes, Len(cRes) - 1) + cAux
               Endif

               nPos = 0
            else
               nPos = nPos + Len(cLetter)

               If nPos >= 73
                  If (Asc(cLetter) == 32) .OR. (Asc(cLetter) == 9)
                     cLetter = "=" + Right(Transform(Asc(cLetter), "@0"), 2)
                  Endif
    
                  cLetter = cLetter + "=" + Chr(13) + Chr(10)
                  nPos = 0
               Endif
            Endif

            cRes = cRes + cLetter

        Next


   Case nMode == 2  && Codifica en Base64, no usado actualmente
        cRes = StrConv(cString, 13)


   Case nMode == 3  && Codifica el cuerpo en Quoted Printable
        cString = StrTran(cString, "\", "\\")
        cString = StrTran(cString, '"', '\"')
   
        cLine = "=?iso-8859-1?Q?"

        For nCont = 1 to Len(cString)

            cLetter = SubStr(cString, nCont, 1)

            Do Case
               Case (cLetter == '"') .OR. (cLetter = "[") .OR. ;
                    (cLetter == "]") .OR. (cLetter = "=") .OR. ;
                    (cLetter == "?") .OR. (cLetter = "_") .OR. ;
                    (Asc(cLetter) == 9)
                    cLetter = "=" + Right(Transform(Asc(cLetter), "@0"), 2)

               Case Asc(cLetter) == 32
                    cLetter = "_"

               Case (Asc(cLetter) >= 33) .AND. (Asc(cLetter) <= 60)
                    * No transformation needed

               Case (Asc(cLetter) >= 62) .AND. (Asc(cLetter) <= 126)
                    * No transformation needed

               OtherWise
                    cLetter = "=" + Right(Transform(Asc(cLetter), "@0"), 2)

            EndCase

            cLine = cLine + cLetter

            If Len(cLine) + nExtraChar >= 71
               cLine = cLine + "?="
               cRes = cRes + cLine + Chr(13) + Chr(10) + " "

               cLine = "=?iso-8859-1?Q?"
               nExtraChar = 1
            Endif

        Next

        cLine = cLine + "?="
        cRes = cRes + cLine

EndCase

Return(cRes)



***
* Function Decode
***
Function Decode(cString, nMode)
Local cRes, nCont, cLetter, cLine, cBefore, cAfter, cDecoded, cEncoding

cRes = ""

Do Case
   Case nMode == 1
        cString = StrTran(cString, "=" + Chr(13) + Chr(10), "")
        cString = StrTran(cString, "=" + Chr(13), "")

        For nCont = 0 to 255
            cLetter = "=" + Right(Transform(nCont, "@0"), 2)
            cString = StrTran(cString, cLetter, Chr(nCont))
        Next

        cRes = cString


   Case nMode == 2
        cRes = StrConv(cString, 14)


   Case nMode == 3
        cRes = ""
        Do While !Empty(cString)

           * Extraer la linea
           If At(Chr(13), cString) == 0
              cLine = cString
              cString = ""
           else
              cLine = Left(cString, At(Chr(13), cString) - 1)
              cString = SubStr(cString, At(Chr(13), cString) + 1)
           Endif


           * Decodificar
           Do While AT("=?", cLine) <> 0
              cBefore = Left(cLine, AT("=?", cLine) - 1)
              cDecoded = SubStr(cLine, AT("=?", cLine) + 2)

              If AT("?", cDecoded, 3) <> 0
                 cEncoding = Upper(SubStr(cDecoded, AT("?", cDecoded) + 1, 1))
                 cDecoded = SubStr(cDecoded, AT("?", cDecoded, 2) + 1)

                 cAfter = SubStr(cDecoded, AT("?=", cDecoded) + 2)
                 cDecoded = Left(cDecoded, AT("?=", cDecoded) - 1)

                 Do Case
                    Case cEncoding == "Q"
                         For nCont = 0 to 255
                             cLetter = "=" + Right(Transform(nCont, "@0"), 2)
                             cDecoded = StrTran(cDecoded, cLetter, Chr(nCont))
                         Next

                         cDecoded = StrTran(cDecoded, "_", " ")

                    Case cEncoding == "B"
                         cDecoded = StrConv(cDecoded, 14)

                    OtherWise
                         cLine = cDecoded

                 EndCase
              else
                 cAfter = ""
                 cDecoded = ""
              Endif

              cLine = cBefore + cDecoded + cAfter
           Enddo


           * Agregar la linea
           cRes = cRes + cLine

        Enddo

        cRes = StrTran(cRes, "\", "")

EndCase

Return(cRes)



***
* Function GetTimeZoneOffset
***
Function GetTimeZoneOffset()
Local cBuffer, nRes, nOffset

cBuffer = Replicate(Chr(0), 172) 
nRes = GetTimeZoneInformation(@cBuffer)

If (nRes == 0) .OR. (nRes == 1) .OR. (nRes == 2)
   nOffset = this.LongToNum(Left(cBuffer, 4)) * -1 / 60
else
   nOffset = 0
Endif

Return(nOffset)



***
* Function GetTimeZoneOffsetMinutes
***
Function GetTimeZoneOffsetMinutes()
Local cBuffer, nRes, nOffset

cBuffer = Replicate(Chr(0), 172) 
nRes = GetTimeZoneInformation(@cBuffer)

If (nRes == 0) .OR. (nRes == 1) .OR. (nRes == 2)
   nOffset = this.LongToNum(Left(cBuffer, 4)) * -1
else
   nOffset = 0
Endif

Return(nOffset)



***
* Function LongToNum
***
Hidden Function LongToNum(cNumber)
Local nRes

Declare RtlMoveMemory in WIN32API as APIMoveMemory ;
        Integer @DestNum, ;
        String @pVoidSource, ;
        Integer nLength

nRes = 0
APIMoveMemory(@nRes, cNumber, 4)

Return(nRes)



***
* Function GetRegValue
***
Function GetRegValue(nRoot, cKey, cValue)
Local lRes, nKey, nType, cBuffer, nLen

lRes = .F.
this.Response = ""

nKey = 0
nType = 0
cBuffer = Space(255)
nLen = Len(cBuffer)

If RegOpenKeyEx(nRoot, cKey, 0, 0x20019, @nKey) == 0

   If RegQueryValueEx(nKey, cValue, 0, @nType, @cBuffer, @nLen) == 0
      this.Response = Left(cBuffer, nLen)
      lRes = .T.
   Endif

   RegCloseKey(nKey)
Endif

Return(lRes)



***
* Function GetINIEntry
***
Function GetINIEntry(cSection, cEntry, cINIFile)
Local cBuffer, nBufSize

cBuffer = Space(2000)
		
If Empty(cINIFile)
   nBufSize = GetWinINI(cSection, cEntry, "", @cBuffer, Len(cBuffer))
else
   nBufSize = GetPrivateINI(cSection, cEntry, "", @cBuffer, Len(cBuffer), cINIFile)
Endif
		
If nBufSize == 0
   Return("")
Endif

Return(Left(cBuffer, nBufSize))



***
* Function GetType
***
Function GetType(cExt)
Local cRes

cExt = Upper(cExt)

Do Case
   Case (cExt = "HTM") .OR. (cExt = "HTML")
        cRes = "text/html"

   Case cExt = "TXT"
        cRes = "text/plain"

   Case cExt = "CSS"
        cRes = "text/css"

   Case cExt = "GIF"
        cRes = "image/gif"

   Case cExt = "PNG"
        cRes = "image/x-png"

   Case (cExt = "JPEG") .OR. (cExt = "JPG")
        cRes = "image/jpeg"

   Case (cExt = "TIFF") .OR. (cExt = "TIF")
        cRes = "image/tiff"

   Case cExt = "RGB"
        cRes = "image/rgb"

   Case cExt = "PICT"
        cRes = "image/x-pict"

   Case cExt = "BMP"
        cRes = "image/x-ms-bmp"

   Case cExt = "PCD"
        cRes = "image/x-photo-cd"

   Case cExt = "FIF"
        cRes = "image/fif"

   Case cExt = "CMX"
        cRes = "image/x-cmx"

   Case cExt = "DWG"
        cRes = "image/x-dwg"

   Case cExt = "DXF"
        cRes = "image/x-dxf"

   Case (cExt = "AU") .OR. (cExt = "SND")
        cRes = "audio/basic"

   Case (cExt = "AIF") .OR. (cExt = "AIFF") .OR. (cExt = "AIFC")
        cRes = "audio/x-aiff"

   Case cExt = "WAV"
        cRes = "audio/x-wav"

   Case (cExt = "RA") .OR. (cExt = "RAM")
        cRes = "application/x-pn-realaudio"

   Case (cExt = "MPEG") .OR. (cExt = "MPG") .OR. (cExt = "MPE")
        cRes = "video/mpeg"

   Case (cExt = "QT") .OR. (cExt = "MOV")
        cRes = "video/quicktime"

   Case cExt = "AVI"
        cRes = "video/x-msvideo"

   Case cExt = "MOVIE"
        cRes = "video/x-sgi-movie"

   Case cExt = "VDO"
        cRes = "video/vdo"

   Case (cExt = "AI") .OR. (cExt = "EPS") .OR. (cExt = "PS")
        cRes = "application/postscript"

   Case cExt = "RTF"
        cRes = "application/rtf"

   Case cExt = "PDF"
        cRes = "application/pdf"

   Case cExt = "TEX"
        cRes = "application/x-tex"

   Case cExt = "GTAR"
        cRes = "application/x-gtar"

   Case cExt = "TAR"
        cRes = "application/x-tar"

   Case cExt = "ZIP"
        cRes = "application/zip"

   Case (cExt = "SIT") .OR. (cExt = "SEA")
        cRes = "application/x-stuffit"

   Case cExt = "JS"
        cRes = "text/javascript"

   Case cExt = "VBS"
        cRes = "text/vbscript"

   Case cExt = "SH"
        cRes = "application/x-sh"

   Case cExt = "CSH"
        cRes = "application/x-csh"

   Case cExt = "PL"
        cRes = "application/x-perl"

   Case cExt = "TCL"
        cRes = "application/x-tcl"

   Case cExt = "PPT"
        cRes = "application/vnd.ms-powerpoint"

   Case cExt = "DOC"
        cRes = "application/msword"

   Case cExt = "XLS"
        cRes = "application/vnd.ms-excel"

   Case cExt = "MDB"
        cRes = "application/msaccess"

   Case cExt = "MA"
        cRes = "application/mathematica"

   OtherWise
        cRes = "application/octet-stream"
EndCase

Return(cRes)


EndDefine
