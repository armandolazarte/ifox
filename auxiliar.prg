
If .F.
   ?  "Creado por Pablo Pioli"
Endif


#include IFOX.H


************************   AUXILIAR FUNCTIONS  *****************************

****************************************************************************


***
* Function IsYes
*   Did you mean Yes?
***
Function IsYes(cValue)

If (cValue == "Sí") .OR. (cValue == "sí") .OR. ;
   (Upper(cValue) == "SI") .OR. (Upper(cValue) == "YES")
   Return(.T.)
else
   Return(.F.)
Endif

EndFunc



***
* Function Delay(nSeconds)
*  Waste your time
***
Function Delay(nSeconds)
Local tEndTime

tEndTime = DateTime() + nSeconds
Do While DateTime() < tEndTime
Enddo

EndFunc



***
* Function Win32Delay(nSeconds)
*  Waste your time
***
Function Win32Delay(nSeconds)

Declare Sleep in Win32API Integer nMillisecs
Sleep(nSeconds * 1000)

EndFunc



***
*  Function Irand(I, J)
*     Generate a random number so I <= X <= J
***
Function Irand(I, J)
Return Int((J - I + 1) * Rand() + I)



***
* Function CopyFileTime
*   Copy the last modification time from a file to another
***
Function CopyFileTime(cFile1, cFile2)
Local tTime, cSysTime, cBuffTime1, cBuffTime2, nHandle

Declare Integer SetFileTime in kernel32 ;
        Integer hFile, ; 
        String lpCreationTime, ; 
        String lpLastAccessTime, ; 
        String lpLastWriteTime 

Declare Integer GetFileAttributesEx in kernel32 ; 
        String  lpFileName, ; 
        Integer fInfoLevelId, ; 
        String  @lpFileInformation 

Declare Integer LocalFileTimeToFileTime in kernel32 ; 
        String LOCALFILETIME, ; 
        String @FILETIME 

Declare Integer FileTimeToSystemTime in kernel32 ; 
        String FILETIME, ; 
        String @SYSTEMTIME 

Declare Integer SystemTimeToFileTime in kernel32 ; 
        String lpSYSTEMTIME, ; 
        String @FILETIME 

Declare Integer _lopen in kernel32 ; 
        String lpFileName, Integer iReadWrite 

Declare Integer _lclose in kernel32 Integer hFile 

tTime = FDate(cFile1, 1)

cSysTime = Int2Word(Year(tTime)) + ; 
           Int2Word(Month(tTime)) + ; 
           Int2Word(Dow(tTime) - 1) + ; 
           Int2Word(Day(tTime)) + ; 
           Int2Word(Hour(tTime)) + ; 
           Int2Word(Minute(tTime)) + ; 
           Int2Word(Sec(tTime)) + ; 
           Int2Word(0) 

cBuffTime1 = Replicate(Chr(0), 16)
SystemTimeToFileTime(cSysTime, @cBuffTime1)

cBuffTime2 = cBuffTime1
LocalFileTimeToFileTime(cBuffTime2, @cBuffTime1)

nHandle = _lopen(cFile2, 2) 

If nHandle < 0 
   Return(.F.)
Endif

SetFileTime(nHandle, cBuffTime1, cBuffTime1, cBuffTime1) 

_lclose(nHandle) 

Return(.T.)


Function Int2Word(nVal)
Return(Chr(Mod(nVal, 256)) + Chr(Int(nVal / 256)))



***
*   Function LastOfMonth(dDate)
*      Get the last day of a month
***
Function LastOfMonth(dDate)
Local cAntCentury, x, nCont

cAntCentury = Set("Century")
Set Century On

x = CtoD("")
nCont = 31

Do While Empty(x)
   x = CtoD(PadL(nCont, 2, "0") + "/" + Right(DtoC(dDate), 7))
   nCont = nCont - 1
Enddo

If cAntCentury == "ON"
   Set Century On
else
   Set Century Off
Endif   

Set Century to 19 RollOver 80

Return(x)



***
* Function GetWinDir()
*   Determina donde se encuentra el directorio WINDOWS
***
Function GetWinDir()
Local cBuffer, nRes, cRes

Declare Integer GetWindowsDirectory in Win32API String @cBuffer, Integer nLen

cBuffer = Space(250)
nRes = GetWindowsDirectory(@cBuffer, Len(cBuffer))

If nRes > 0
   cRes = RTrim(cBuffer)

   If Right(cRes, 1) == Chr(0)
      cRes = Left(cRes, Len(cRes) - 1)
   Endif

   If Right(cRes, 1) <> "\"
      cRes = cRes + "\"
   Endif
else
   cRes = ""
Endif

Return(cRes)



***
* Function GetTempDir()
*   Determina donde se encuentra el directorio temporal
***
Function GetTempDir()
Local cBuffer, nRes, cRes

Declare Long GetTempPath in Win32API Long, String @

cBuffer = Space(250)
nRes = GetTempPath(Len(cBuffer), @cBuffer)

If nRes > 0
   cRes = RTrim(cBuffer)

   If Right(cRes, 1) == Chr(0)
      cRes = Left(cRes, Len(cRes) - 1)
   Endif

   If Right(cRes, 1) <> "\"
      cRes = cRes + "\"
   Endif
else
   cRes = ""
Endif

Return(cRes)



***
*   Function TmpName(nLen)
*   Retorna un nombre unico 
***
Function TmpName(nLen)
Local nCont, cNom_Arch

If Type("nLen") <> "N"
   nLen = 8
Endif

cNom_Arch = ""
For nCont = 1 TO nLen
    cNom_Arch = cNom_Arch + Chr(Irand(65, 90))
Next

Return(cNom_Arch)


***
* WinAPI Functions
***

Function APINumtoStr(nNum, nLength)
local cRes, nCont, nTmp, lnMax

nMax = (256 ^ nLength) - 1

If nNum < 0
   nNum = (2 ^ (nLength * 8)) + nNum
Endif

nNum = BitAND(nNum, nMax)

cRes = ""	
For nCont = (nLength - 1) to 0 Step -1
    nTmp = Int(nNum / 256 ^ nCont)
    nNum = nNum - nTmp * (256 ^ nCont)
    cRes = Chr(nTmp) + cRes
Next

Return(cRes)



Function APIStrtoNum(cNumber)
local nRes, nCont, nLength, cTmp, nPower

nRes = 0
nLength = Len(cNumber)
nPower = 1

For nCont = 1 to nLength
    nTmp = Asc(SubStr(cNumber, nCont, 1))
    nRes = nRes + (nTmp * nPower)
    nPower = nPower * 256
Next

Return(nRes)



***
* Function DectoBin(nValor)
*   Convierte un numero de Decimal a Binario
***
Function DectoBin(nValor)
Local cRes, nResto

cRes = ""
Do While nValor <> 0
   nResto = (nValor / 2) - Int(nValor / 2)
   cRes = PadL(nResto * 2, 1, "0") + cRes
   nValor = Int(nValor / 2)
Enddo

Return(cRes)



***
* Function BintoDec(cValor)
*   Convierte un numero de Binario a Decimal
***
Function BintoDec(cValor)
Local nRes, nCont, cLetter, nBase

If Empty(cValor)
   Return(0)
Endif

nRes = 0
nBase = 1
For nCont = Len(cValor) to 1 Step -1
    cLetter = SubStr(cValor, nCont, 1)
    nRes = nRes + (Val(cLetter) * nBase)
    nBase = nBase * 2
Next

Return(nRes)



***
* Function HextoDec(cValor)
*   Convierte un numero de Hexadecimal a Decimal
***
Function HextoDec(cValor)
Local nRes, aEquivalencias, nCont, cLetter, nPos, nBase

If Empty(cValor)
   Return(0)
Endif

Dimension aEquivalencias[16]
aEquivalencias[ 1] = "0"
aEquivalencias[ 2] = "1"
aEquivalencias[ 3] = "2"
aEquivalencias[ 4] = "3"
aEquivalencias[ 5] = "4"
aEquivalencias[ 6] = "5"
aEquivalencias[ 7] = "6"
aEquivalencias[ 8] = "7"
aEquivalencias[ 9] = "8"
aEquivalencias[10] = "9"
aEquivalencias[11] = "A"
aEquivalencias[12] = "B"
aEquivalencias[13] = "C"
aEquivalencias[14] = "D"
aEquivalencias[15] = "E"
aEquivalencias[16] = "F"

cValor = Upper(cValor)

nRes = 0
nBase = 1
For nCont = Len(cValor) to 1 Step -1
    cLetter = SubStr(cValor, nCont, 1)
    nPos = AScan(aEquivalencias, cLetter)

    If nPos == 0
       Return(0)
    Endif

    nRes = nRes + ((nPos - 1) * nBase)
    nBase = nBase * 16
Next

Return(nRes)
