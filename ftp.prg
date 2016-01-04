
If .F.
   ?  "Creado por Pablo Pioli"
Endif


#include IFOX.H


#define INTERNET_DEFAULT_FTP_PORT           21
#define INTERNET_OPEN_TYPE_DIRECT            1
#define INTERNET_SERVICE_FTP	             1
#define FTP_TRANSFER_TYPE_BINARY             2
#define INTERNET_FLAG_RELOAD        2147483648
#define GENERIC_READ                0x80000000
#define GENERIC_WRITE               0x40000000


#define BYTE_1	                             1
#define BYTE_2                             256
#define BYTE_3                           65536
#define BYTE_4                        16777216
#define MAXDWORD	                4294967295

#define BIT_ATTRIBUTE_READONLY        	 	 0
#define BIT_ATTRIBUTE_HIDDEN          	 	 1
#define BIT_ATTRIBUTE_SYSTEM          		 2
#define BIT_ATTRIBUTE_DIRECTORY              4
#define BIT_ATTRIBUTE_ARCHIVE        	 	 5
#define BIT_ATTRIBUTE_ENCRYPTED		         6
#define BIT_ATTRIBUTE_NORMAL         	 	 7
#define BIT_ATTRIBUTE_TEMPORARY      	 	 8
#define BIT_ATTRIBUTE_SPARSE_FILE	         9
#define BIT_ATTRIBUTE_REPARSE_POINT         10
#define BIT_ATTRIBUTE_COMPRESSED            11
#define BIT_ATTRIBUTE_OFFLINE               12

#define ERROR_NO_MORE_FILES	                18

#define MAX_PATH                           260
#define iNULL                            Chr(0)


* Directory
#define DIR_FileName                  1
#define DIR_Alternate_FileName        2
#define DIR_File_Size                 3
#define DIR_File_Create_Date          4
#define DIR_File_Last_Access_Time     5
#define DIR_File_Last_Write_Time      6
#define DIR_File_Attributes           7



***************************   FTP CLASS   **********************************

****************************************************************************
Define Class FTP as iFox of iFox.PRG OlePublic

             Hidden Server
             Hidden UserName
             Hidden UserPassword
             Hidden Port

             Hidden InetHandle
             Hidden FTPHandle

             ErrorNumber = 0
             ErrorMessage = 0
             ExtendedErrorNumber = 0

             Dimension DirFiles[1]

             Hidden FileHandle
             Hidden InternetFileHandle
             EOT = .F.
             TransferredBytes = 0


***
* Function eInit()
***
Function eInit()


* Initialize properties
this.Server = ""
this.UserName = ""
this.UserPassword = ""
this.Port = 0

this.InetHandle = 0
this.FTPHandle = 0

this.FileHandle = 0
this.InternetFileHandle = 0


* Declare Kernel32 Functions
Declare Integer GetLastError in Kernel32
Declare Integer FileTimeToSystemTime in Kernel32 String @lpcBuffer, String @lpcBuffer

		   
* Declare Internet Connection Functions
Declare Integer InternetOpen in WinInet String @lpcAgent, Integer nAccessType, String @lpcProxyName, String @lpcProxyBypass, Integer nFlags
Declare Integer InternetConnect in WinInet Integer nHandle, String @lpcServer, Short nPort, String @lpcUserName, String @lpcPassword, Integer nService, Integer nFlags, Integer nContext
Declare Integer InternetCloseHandle in WinInet Integer nHandle

		   			
* Declare FTP Functions
Declare Integer FtpCreateDirectory in WinInet Integer nHandle, String @lpcDirectory
Declare Integer FtpDeleteFile in WinInet Integer nHandle, String @lpcFileName
Declare Integer FtpFindFirstFile in WinInet Integer nHandle, String @lpcSearchStr, String @lpcWIN32_FIND_DATA, Integer nFlags, Integer nContext
Declare Integer InternetFindNextFile in WinInet Integer nHandle, String @lpcWIN32_FIND_DATA
Declare Integer FtpGetCurrentDirectory in WinInet Integer nHandle, String @lpcDirectory, Integer @nMaxPath
Declare Integer FtpGetFile in WinInet Integer nHandle, String @lpcRemoteFile, String @lpcNewFile, Integer nFailIfExists, Integer nAttributes, Integer nFlags, Integer nContext
Declare Integer FtpOpenFile in WinInet Integer nHandle, String lpcRemoteFile, Integer nAccessType, Integer nFlags, Integer nContext
Declare Integer FtpPutFile in WinInet Integer nHandle, String @lpcNewFile, String @lpcRemoteFile, Integer nFlags, Integer nContext
Declare Integer FtpRemoveDirectory in WinInet Integer nHandle, String @lpcDirectory
Declare Integer FtpRenameFile in WinInet Integer nHandle, String @lpcRemoteFile, String @lpcNewFile
Declare Integer FtpSetCurrentDirectory in WinInet Integer nHandle, String @lpcDirectory
Declare Integer InternetGetLastResponseInfo in WinInet Integer @nError, String @lpcBuffer, Integer @nMaxPath
Declare Integer InternetReadFile in WinInet Integer nHandle, String @cBuffer, Long nBytestoRead, Long @nBytesRead
Declare Integer InternetWriteFile in WinInet Integer nHandle, String cBuffer, Long nBytestoWrite, Long @nBytesWritten

		   			
EndFunc



***
* Function Connect
*   Stablish a connection with an FTP server
***
Function Connect(cServer, cUserName, cPassword, nPort)


* Validate the port
If Type("nPort") <> "N"
   nPort = INTERNET_DEFAULT_FTP_PORT
Endif


* Open an Internet connection
this.InetHandle = InternetOpen("", INTERNET_OPEN_TYPE_DIRECT, "", "", 0)

If this.InetHandle == 0
   this.GetExtendedError()
   Return(.F.)
Endif


* Connect to the FTP server
this.FTPHandle = InternetConnect(this.InetHandle, cServer, INTERNET_DEFAULT_FTP_PORT, cUserName, cPassword, INTERNET_SERVICE_FTP, 0, 0)
   
If this.FTPHandle == 0
   this.GetExtendedError()

   If this.ErrorNumber == 12014
      this.ErrorMessage = "Nombre de usuario o clave de acceso inválida"
   Endif

   InternetCloseHandle(this.InetHandle)
   Return(.F.)
Endif

this.Server = cServer
this.UserName = cUserName
this.UserPassword = cPassword
this.Port = nPort

Return(.T.)



***
* Function Close()
*   Cierra la conexion establecida con Connect
***
Function Close()

InternetCloseHandle(this.FTPHandle)
InternetCloseHandle(this.InetHandle)

EndFunc



***
* Function GetExtendedError()
*   Get a description for the last error
***
Hidden Function GetExtendedError()
Local nError, cBuffer

this.ErrorNumber = GetLastError()

nError = 0
cBuffer = Space(MAX_PATH)
	        
InternetGetLastResponseInfo(nError, @cBuffer, MAX_PATH)

this.ExtendedErrorNumber = nError
this.ErrorMessage = Left(cBuffer, At(iNULL, cBuffer) - 1)

EndFunc



***
* Function CD(cDir)
*   Changes the current directory
***
Function CD(cDir)
Local lRes

If FtpSetCurrentDirectory(this.FTPHandle, cDir) == 1
   lRes = .T.
else
   lRes = .F.
Endif

Return(lRes)



***
* Function Upload
*   Uploads a file
***
Function Upload(cSource, cDestination)
Local lRes

If FTPPutFile(this.FTPHandle, cSource, cDestination, FTP_TRANSFER_TYPE_BINARY, 0) == 1
   lRes = .T.
else
   lRes = .F.
Endif

Return(lRes)



***
* Function Download
*   Downloads a file
***
Function Download(cSource, cDestination)
Local lRes

If FTPGetFile(this.FTPHandle, cSource, cDestination, .T., INTERNET_FLAG_RELOAD, FTP_TRANSFER_TYPE_BINARY, 0) == 1
   lRes = .T.
else
   lRes = .F.
Endif

Return(lRes)



***
* Function StartDownload
*   Start a download
***
Function StartDownload(cSource, cDestination)


* Inicializar propiedades
this.EOT = .F.
this.TransferredBytes = 0
this.InternetFileHandle = 0
this.FileHandle = -1


* Abrir archivo de origen
this.InternetFileHandle = FtpOpenFile(this.FTPHandle, cSource, GENERIC_READ, FTP_TRANSFER_TYPE_BINARY + INTERNET_FLAG_RELOAD, 0)
If this.InternetFileHandle == 0
   this.ErrorNumber = -101
   Return(.F.)
Endif


* Abrir archivo de destino
this.FileHandle = FCreate(cDestination)

If this.FileHandle == -1
   this.ErrorNumber = -102
   Return(.F.)
Endif


Return(.T.)



***
* Function DownloadNextPart
*   Continua una descarga
***
Function DownloadNextPart(nBytes)
Local cBuffer, nRead

If Type("nBytes") <> "N"
   nBytes = 4096
Endif

cBuffer = Space(nBytes)
nRead = 0
If InternetReadFile(this.InternetFileHandle, @cBuffer, Len(cBuffer), @nRead) <> 0

   If nRead == 0
      this.EOT = .T.
   else
      If FWrite(this.FileHandle, Left(cBuffer, nRead)) <> nRead
         this.ErrorNumber = -103
         Return(.F.)
      Endif

      this.TransferredBytes = this.TransferredBytes + nRead
   Endif
else
   this.ErrorNumber = -104
   Return(.F.)
Endif

Return(.T.)



***
* Function EndDownload
*   Finaliza un download
***
Function EndDownload()

If this.InternetFileHandle <> 0
   InternetCloseHandle(this.InternetFileHandle)
   this.InternetFileHandle = 0
Endif

If this.FileHandle <> -1
   FClose(this.FileHandle)
   this.FileHandle = -1
Endif

EndFunc



***
* Function StartUpload
*   Comienza un upload
***
Function StartUpload(cSource, cDestination)


* Inicializar propiedades
this.EOT = .F.
this.TransferredBytes = 0
this.InternetFileHandle = 0
this.FileHandle = -1


* Abrir archivo de origen
this.InternetFileHandle = FtpOpenFile(this.FTPHandle, cDestination, GENERIC_WRITE, FTP_TRANSFER_TYPE_BINARY + INTERNET_FLAG_RELOAD, 0)
If this.InternetFileHandle == 0
   this.ErrorNumber = -102
   Return(.F.)
Endif


* Abrir archivo de destino
this.FileHandle = FOpen(cSource)

If this.FileHandle == -1
   this.ErrorNumber = -101
   Return(.F.)
Endif


Return(.T.)



***
* Function UploadNextPart
*   Continua un upload
***
Function UploadNextPart(nBytes)
Local cData, nWritten

If Type("nBytes") <> "N"
   nBytes = 4096
Endif

cData = FRead(this.FileHandle, nBytes)

If Len(cData) <> 0
   nWritten = 0
   If InternetWriteFile(this.InternetFileHandle, @cData, ;
                        Len(cData), @nWritten) <> 0

      If nWritten <> Len(cData)
         this.ErrorNumber = -105
         Return(.F.)
      Endif

      this.TransferredBytes = this.TransferredBytes + nWritten
   else
      this.ErrorNumber = -104
      Return(.F.)
   Endif
Endif

If FEOF(this.FileHandle)
   this.EOT = .T.
Endif

Return(.T.)



***
* Function EndUpload
*   Finaliza un upload
***
Function EndUpload()

If this.InternetFileHandle <> 0
   InternetCloseHandle(this.InternetFileHandle)
   this.InternetFileHandle = 0
Endif

If this.FileHandle <> -1
   FClose(this.FileHandle)
   this.FileHandle = -1
Endif

EndFunc



***
* Function CreateFolder
*   Creates a folder
***
Function CreateFolder(cFolder)
Local lRes

If FTPCreateDirectory(this.FTPHandle, cFolder) == 1
   lRes = .T.
else
   lRes = .F.
Endif

Return(lRes)



***
* Function DeleteFolder
*   Deletes a folder
***
Function DeleteFolder(cFolder)
Local lRes

If FTPRemoveDirectory(this.FTPHandle, cDir) == 1
   lRes = .T.
else
   lRes = .F.
Endif

Return(lRes)



***
* Function DeleteFile
*   Deletes a file
***
Function DeleteFile(cFile)
Local lRes

If FTPDeleteFile(this.FTPHandle, cFile) == 1
   lRes = .T.
else
   lRes = .F.
Endif

Return(lRes)



***
* Function RenameFile
*   Renames a file
***
Function RenameFile(cOldName, cNewName)
Local lRes

If FTPRenameFile(this.FTPHandle, cOldName, cNewName) == 1
   lRes = .T.
else
   lRes = .F.
Endif

Return(lRes)



***
* Function DIR
*   Return the contents of a folder
***
Function DIR(cFolder, cMask)
Local aFiles, nRes, nCont

Dimension aFiles[1]
nRes = this.VFPScriptDir(cFolder, cMask, @aFiles)

If nRes > 0
   Dimension this.DirFiles[nRes]
   
   For nCont = 1 to nRes
       this.DirFiles[nCont] = CreateObject("FTPFile")

       this.DirFiles[nCont].FileName = aFiles[nCont, 1]
       this.DirFiles[nCont].AlternateName = aFiles[nCont, 2]
       this.DirFiles[nCont].Size = aFiles[nCont, 3]
       this.DirFiles[nCont].CreateDate = aFiles[nCont, 4]
       this.DirFiles[nCont].LastAccessTime = aFiles[nCont, 5]
       this.DirFiles[nCont].LastWriteTime = aFiles[nCont, 6]
       this.DirFiles[nCont].Attributes = aFiles[nCont, 7]
   Next
Endif

Return(nRes)



***
* Function VFPScriptDIR
*   Return the contents of a folder
***
Function VFPScriptDIR(cFolder, cMask, aRes)
Local lRes, cStruct, nFirstFileHandle, nResult


* We need to reconnect to read another folder
this.Close()
If !this.Connect(this.Server, this.UserName, this.UserPassword, this.Port)
   Return(-2)
Endif

If !Empty(cFolder)
   If !this.CD(cFolder)
      Return(-3)
   Endif
Endif


* Process
Dimension aRes[1, 7]
aRes[1, 1] = .F.

cMask = cMask + iNULL

cStruct = Space(319)

nFirstFileHandle = FtpFindFirstFile(this.FTPHandle, @cMask, @cStruct, INTERNET_FLAG_RELOAD, 0)
this.GetExtendedError()

If nFirstFileHandle == 0
   Return(-1)
Endif

If this.ExtendedErrorNumber == ERROR_NO_MORE_FILES
   Return(0)
Endif

this.ProcessFile(cStruct, @aRes)

Do While .T.

   cStruct = Space(319)
   nResult = InternetFindNextFile(nFirstFileHandle, @cStruct)
   this.GetExtendedError()

   If (nResult == 0) .OR. (this.ExtendedErrorNumber == ERROR_NO_MORE_FILES)
      Exit
   Endif

   this.ProcessFile(cStruct, @aRes)

Enddo

Return(ALen(aRes, 1))



***
* ProcessFile
*   Auxiliar function
***
Hidden Function ProcessFile(cString, aRes)
Local lcFileName, lcAlterName, lnSizeHigh, lnSizeLow, lnFileSize
Local lcAttributes, lnArrayLen, lcTimeBuff, ldCreateDate
Local ldAccessDate, ldWriteDate, laNewArray, lnResult

If Type("aRes[1, 1]") == "L"
   Dimension aRes [1, 7]
else	
   Dimension aRes[ALen(aRes, 1) + 1, 7]
Endif
     
lnArrayLen = ALen(aRes, 1)
          		       	      
lcFileName = SubStr(cString, 45, MAX_PATH)
lcAlterName = Right(cString, 14)
	      	
lcFileName = Left(lcFileName, AT(iNull, lcFileName) - 1)
lcAlterName = Left(lcAlterName, AT(iNull, lcAlterName) - 1)
	      	
lnSizeHigh = (Asc(SubStr(cString, 29, 1)) * BYTE_1) + ;
             (Asc(SubStr(cString, 30, 1)) * BYTE_2) + ;
	      	 (Asc(SubStr(cString, 31, 1)) * BYTE_3) + ;
	      	 (Asc(SubStr(cString, 32, 1)) * BYTE_4) 
	      				 
lnSizeLow =  (Asc(SubStr(cString, 33, 1)) * BYTE_1) + ;
			 (Asc(SubStr(cString, 34, 1)) * BYTE_2) + ;
	    	 (Asc(SubStr(cString, 35, 1)) * BYTE_3) + ;
	    	 (Asc(SubStr(cString, 36, 1)) * BYTE_4) 
	      				 
lnFileSize = (lnSizeHigh * MAXDWORD) + lnSizeLow
	      	
lcTimeBuff = SubStr(cString, 5, 8)
ldCreateDate = this.ProcessDate(lcTimeBuff)
	        
lcTimeBuff = SubStr(cString, 13, 8)
ldAccessDate = this.ProcessDate(lcTimeBuff)
	        
lcTimeBuff = SubStr(cString, 21, 8)
ldWriteDate = this.ProcessDate(lcTimeBuff)
	        
lcAttributes = this.ProcessAttributes(Left(cString, 4))
	        
aRes[lnArrayLen, 1] = AllTrim(lcFileName)
aRes[lnArrayLen, 2] = AllTrim(lcAlterName)
aRes[lnArrayLen, 3] = lnFileSize
aRes[lnArrayLen, 4] = ldCreateDate
aRes[lnArrayLen, 5] = ldAccessDate
aRes[lnArrayLen, 6] = ldWriteDate
aRes[lnArrayLen, 7] = lcAttributes
  	 
EndFunc



***
* Function ProcessDate
*   Auxiliar function
***
Hidden Function ProcessDate(cBuffer)
Local lcInBuffer, ldDateTime, fResult, lcBuild
Local lnDay, lnMonth, lnYear, lnHour, lnMinute, lnSecond
		
lcInBuffer = Space(16)
		
fResult = FileTimeToSystemTime(@cBuffer, @lcInBuffer)
this.GetExtendedError()

If fResult = 0   && Failed
   ldDateTime = {^1901/01/01 00:00:01}   && Default Time
   Return(ldDateTime)
Endif
			
lnYear = Asc(SubStr(lcInBuffer, 1, 1)) + (Asc(SubStr(lcInBuffer, 2, 1)) * BYTE_2)
lnMonth = Asc(SubStr(lcInBuffer, 3, 1)) + (Asc(SubStr(lcInBuffer, 4, 1)) * BYTE_2)
lnDay = Asc(SubStr(lcInBuffer, 7, 1)) + (Asc(SubStr(lcInBuffer, 8, 1)) * BYTE_2)
lnHour = Asc(SubStr(lcInBuffer, 9, 1)) + (Asc(SubStr(lcInBuffer, 10, 1)) * BYTE_2)
lnMinute = Asc(SubStr(lcInBuffer, 11, 1)) + (Asc(SubStr(lcInBuffer, 12, 1)) * BYTE_2)
lnSecond = Asc(SubStr(lcInBuffer, 13, 1)) + (Asc(SubStr(lcInBuffer, 13, 1)) * BYTE_2)
			
lcBuild = "^" + AllTrim(Str(lnYear)) + '/' + AllTrim(Str(lnMonth)) + '/' + AllTrim(Str(lnDay)) + ' ' + ;
		  AllTrim(Str(lnHour)) + ':' + AllTrim(Str(lnMinute)) + ':' + AllTrim(Str(lnSecond))
				      
ldDateTime = {&lcBuild}
						  
Return(ldDateTime)
		   


***
* Function ProcessAttributes
*   Auxiliar function
***
Hidden Function ProcessAttributes(cBuffer)
Local cAttributes, lnValue

cAttributes = ""
	 		 		
lnValue = (Asc(SubStr(cBuffer, 1, 1)) * BYTE_1) + ;
		  (Asc(SubStr(cBuffer, 2, 1)) * BYTE_2) + ;
		  (Asc(SubStr(cBuffer, 3, 1)) * BYTE_3) + ;
		  (Asc(SubStr(cBuffer, 4, 1)) * BYTE_4) 

Do Case
   Case BitTest(lnValue, BIT_ATTRIBUTE_READONLY) 
        cAttributes = cAttributes + "R"
   Case BitTest(lnValue, BIT_ATTRIBUTE_HIDDEN) 
        cAttributes = cAttributes + "H"
   Case BitTest(lnValue, BIT_ATTRIBUTE_SYSTEM) 
        cAttributes = cAttributes + "S"
   Case BitTest(lnValue, BIT_ATTRIBUTE_DIRECTORY) 
        cAttributes = cAttributes + "D"
   Case BitTest(lnValue, BIT_ATTRIBUTE_ARCHIVE) 
        cAttributes = cAttributes + "A"
   Case BitTest(lnValue, BIT_ATTRIBUTE_NORMAL) 
        cAttributes = cAttributes + "N"
   Case BitTest(lnValue, BIT_ATTRIBUTE_TEMPORARY) 
        cAttributes = cAttributes + "T"
   Case BitTest(lnValue, BIT_ATTRIBUTE_COMPRESSED) 
        cAttributes = cAttributes + "C"
   Case BitTest(lnValue, BIT_ATTRIBUTE_OFFLINE) 
        cAttributes = cAttributes + "O"
EndCase

Return(cAttributes)


EndDefine



**************************   FTPFILE CLASS   *******************************

****************************************************************************
Define Class FTPFile as Custom

             FileName = ""
             AlternateName = ""
             Size = 0
             CreateDate = CtoT("")
             LastAccessTime = CtoT("")
             LastWriteTime = CtoT("")
             Attributes = ""

EndDefine
