
If .F.
   ?  "Creado por Pablo Pioli"
Endif


#include IFOX.H

#define NEW_LINE  Chr(13) + Chr(10)


*****************************   XML CLASS   ********************************

****************************************************************************
Define Class XML as Custom OlePublic

       Parser = "Msxml2.DOMDocument.4.0"
       XSDType = XSD_INLINE

       Status = .F.
       Interval = 100

       Dimension Tables[1]
       Hidden Handle

       Hidden TempXML
       Hidden TempItemName


***
*  Function AddTable(cTable)
***
Function AddTable(cTable)

Dimension this.Tables[ALen(this.Tables) + 1]
this.Tables[ALen(this.Tables)] = cTable

Return(cTable)



***
*  Function CursortoXML(cFile)
***
Function CursortoXML(cFile)
Local cRes, cAntPoint, cAntSeparator
Local nCont, aCampos, nFields, nPos, cItem, xValue

cRes = ""

cAntPoint = Set("Point")
cAntSeparator = Set("Separator")

If Empty(cFile)
   cRes = '<?xml version = "1.0" encoding="Windows-1252" standalone="yes"?>' + NEW_LINE
   cRes = cRes + "<VFPData>" + NEW_LINE
else
   this.Handle = FCreate(cFile)
   If this.Handle == -1
      Return(.F.)
   Endif

   If !this.Write('<?xml version = "1.0" encoding="Windows-1252" standalone="yes"?>')
      Return(.F.)
   Endif

   If !this.Write("<VFPData>")
      Return(.F.)
   Endif
Endif



* Generar Schema
If this.XSDType == XSD_INLINE

If Empty(cFile)
   cRes = cRes + Space(3) + '<xsd:schema id="VFPData" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:msdata="urn:schemas-microsoft-com:xml-msdata">' + NEW_LINE
   cRes = cRes + Space(6) + '<xsd:element name="VFPData" msdata:IsDataSet="true">' + NEW_LINE
   cRes = cRes + Space(9) + '<xsd:complexType>' + NEW_LINE
   cRes = cRes + Space(12) + '<xsd:choice maxOccurs="unbounded">' + NEW_LINE
else
   If !this.Write(Space(3) + '<xsd:schema id="VFPData" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:msdata="urn:schemas-microsoft-com:xml-msdata">')
      Return(.F.)
   Endif

   If !this.Write(Space(6) + '<xsd:element name="VFPData" msdata:IsDataSet="true">')
      Return(.F.)
   Endif

   If !this.Write(Space(9) + '<xsd:complexType>')
      Return(.F.)
   Endif

   If !this.Write(Space(12) + '<xsd:choice maxOccurs="unbounded">')
      Return(.F.)
   Endif
Endif


For nCont = 2 to ALen(this.Tables)

    If Empty(cFile)
       cRes = cRes + Space(15) + '<xsd:element name="' + this.Tables[nCont] + '">' + NEW_LINE
       cRes = cRes + Space(18) + '<xsd:complexType>' + NEW_LINE
       cRes = cRes + Space(21) + '<xsd:sequence>' + NEW_LINE
    else
       If !this.Write(Space(15) + '<xsd:element name="' + this.Tables[nCont] + '">')
          Return(.F.)
       Endif

       If !this.Write(Space(18) + '<xsd:complexType>')
          Return(.F.)
       Endif

       If !this.Write(Space(21) + '<xsd:sequence>')
          Return(.F.)
       Endif
    Endif

    Select(this.Tables[nCont])

    Dimension aCampos[1]
    nFields = AFields(aCampos)

    For nPos = 1 to nFields

        If Empty(cFile)
           Do Case
              Case aCampos[nPos, 2] == "C"
                   cRes = cRes + Space(24) + '<xsd:element name="' + Lower(aCampos[nPos, 1]) + '">' + NEW_LINE
                   cRes = cRes + Space(27) + '<xsd:simpleType>' + NEW_LINE
                   cRes = cRes + Space(30) + '<xsd:restriction base="xsd:string">' + NEW_LINE
                   cRes = cRes + Space(33) + '<xsd:maxLength value="' + LTrim(Str(aCampos[nPos, 3], 3, 0)) +'"/>' + NEW_LINE
                   cRes = cRes + Space(30) + '</xsd:restriction>' + NEW_LINE
                   cRes = cRes + Space(27) + '</xsd:simpleType>' + NEW_LINE
                   cRes = cRes + Space(24) + '</xsd:element>' + NEW_LINE


              Case aCampos[nPos, 2] == "D"
                   cRes = cRes + Space(24) + '<xsd:element name="' + Lower(aCampos[nPos, 1]) + '" type="xsd:date"/>' + NEW_LINE


              Case aCampos[nPos, 2] == "L"
                   cRes = cRes + Space(24) + '<xsd:element name="' + Lower(aCampos[nPos, 1]) + '" type="xsd:boolean"/>' + NEW_LINE


              Case aCampos[nPos, 2] == "M"
                   cRes = cRes + Space(24) + '<xsd:element name="' + Lower(aCampos[nPos, 1]) + '">' + NEW_LINE
                   cRes = cRes + Space(27) + '<xsd:simpleType>' + NEW_LINE
                   cRes = cRes + Space(30) + '<xsd:restriction base="xsd:string">' + NEW_LINE
                   cRes = cRes + Space(33) + '<xsd:maxLength value="2147483647"/>' + NEW_LINE
                   cRes = cRes + Space(30) + '</xsd:restriction>' + NEW_LINE
                   cRes = cRes + Space(27) + '</xsd:simpleType>' + NEW_LINE
                   cRes = cRes + Space(24) + '</xsd:element>' + NEW_LINE


              Case aCampos[nPos, 2] == "N"
                   cRes = cRes + Space(24) + '<xsd:element name="' + Lower(aCampos[nPos, 1]) + '">' + NEW_LINE
                   cRes = cRes + Space(27) + '<xsd:simpleType>' + NEW_LINE
                   cRes = cRes + Space(30) + '<xsd:restriction base="xsd:decimal">' + NEW_LINE
                   cRes = cRes + Space(33) + '<xsd:totalDigits value="' + LTrim(Str(aCampos[nPos, 3], 3, 0)) +'"/>' + NEW_LINE
                   cRes = cRes + Space(33) + '<xsd:fractionDigits value="' + LTrim(Str(aCampos[nPos, 4], 3, 0)) +'"/>' + NEW_LINE
                   cRes = cRes + Space(30) + '</xsd:restriction>' + NEW_LINE
                   cRes = cRes + Space(27) + '</xsd:simpleType>' + NEW_LINE
                   cRes = cRes + Space(24) + '</xsd:element>' + NEW_LINE


              Case aCampos[nPos, 2] == "F"
                   cRes = cRes + Space(24) + '<xsd:element name="' + Lower(aCampos[nPos, 1]) + '">' + NEW_LINE
                   cRes = cRes + Space(27) + '<xsd:simpleType>' + NEW_LINE
                   cRes = cRes + Space(30) + '<xsd:restriction base="xsd:decimal">' + NEW_LINE
                   cRes = cRes + Space(33) + '<xsd:totalDigits value="' + LTrim(Str(aCampos[nPos, 3], 3, 0)) +'"/>' + NEW_LINE
                   cRes = cRes + Space(33) + '<xsd:fractionDigits value="' + LTrim(Str(aCampos[nPos, 4], 3, 0)) +'"/>' + NEW_LINE
                   cRes = cRes + Space(30) + '</xsd:restriction>' + NEW_LINE
                   cRes = cRes + Space(27) + '</xsd:simpleType>' + NEW_LINE
                   cRes = cRes + Space(24) + '</xsd:element>' + NEW_LINE


              Case aCampos[nPos, 2] == "I"
                   cRes = cRes + Space(24) + '<xsd:element name="' + Lower(aCampos[nPos, 1]) + '" type="xsd:int"/>' + NEW_LINE


              Case aCampos[nPos, 2] == "B"
                   cRes = cRes + Space(24) + '<xsd:element name="' + Lower(aCampos[nPos, 1]) + '" type="xsd:double"/>' + NEW_LINE


              Case aCampos[nPos, 2] == "Y"
                   cRes = cRes + Space(24) + '<xsd:element name="' + Lower(aCampos[nPos, 1]) + '">' + NEW_LINE
                   cRes = cRes + Space(27) + '<xsd:simpleType>' + NEW_LINE
                   cRes = cRes + Space(30) + '<xsd:restriction base="xsd:decimal">' + NEW_LINE
                   cRes = cRes + Space(33) + '<xsd:totalDigits value="19"/>' + NEW_LINE
                   cRes = cRes + Space(33) + '<xsd:fractionDigits value="4"/>' + NEW_LINE
                   cRes = cRes + Space(30) + '</xsd:restriction>' + NEW_LINE
                   cRes = cRes + Space(27) + '</xsd:simpleType>' + NEW_LINE
                   cRes = cRes + Space(24) + '</xsd:element>' + NEW_LINE


              Case aCampos[nPos, 2] == "T"
                   cRes = cRes + Space(24) + '<xsd:element name="' + Lower(aCampos[nPos, 1]) + '" type="xsd:dateTime"/>' + NEW_LINE

           EndCase
        else
           Do Case
              Case aCampos[nPos, 2] == "C"
                   If !this.Write(Space(24) + '<xsd:element name="' + Lower(aCampos[nPos, 1]) + '">')
                      Return(.F.)
                   Endif

                   If !this.Write(Space(27) + '<xsd:simpleType>')
                      Return(.F.)
                   Endif

                   If !this.Write(Space(30) + '<xsd:restriction base="xsd:string">')
                      Return(.F.)
                   Endif

                   If !this.Write(Space(33) + '<xsd:maxLength value="' + LTrim(Str(aCampos[nPos, 3], 3, 0)) +'"/>')
                      Return(.F.)
                   Endif

                   If !this.Write(Space(30) + '</xsd:restriction>')
                      Return(.F.)
                   Endif
 
                   If !this.Write(Space(27) + '</xsd:simpleType>')
                      Return(.F.)
                   Endif

                   If !this.Write(Space(24) + '</xsd:element>')
                      Return(.F.)
                   Endif


              Case aCampos[nPos, 2] == "D"
                   If !this.Write(Space(24) + '<xsd:element name="' + Lower(aCampos[nPos, 1]) + '" type="xsd:date"/>')
                      Return(.F.)
                   Endif


              Case aCampos[nPos, 2] == "L"
                   If !this.Write(Space(24) + '<xsd:element name="' + Lower(aCampos[nPos, 1]) + '" type="xsd:boolean"/>')
                      Return(.F.)
                   Endif


              Case aCampos[nPos, 2] == "M"
                   If !this.Write(Space(24) + '<xsd:element name="' + Lower(aCampos[nPos, 1]) + '">')
                      Return(.F.)
                   Endif

                   If !this.Write(Space(27) + '<xsd:simpleType>')
                      Return(.F.)
                   Endif

                   If !this.Write(Space(30) + '<xsd:restriction base="xsd:string">')
                      Return(.F.)
                   Endif

                   If !this.Write(Space(33) + '<xsd:maxLength value="2147483647"/>')
                      Return(.F.)
                   Endif

                   If !this.Write(Space(30) + '</xsd:restriction>')
                      Return(.F.)
                   Endif
 
                   If !this.Write(Space(27) + '</xsd:simpleType>')
                      Return(.F.)
                   Endif

                   If !this.Write(Space(24) + '</xsd:element>')
                      Return(.F.)
                   Endif


              Case aCampos[nPos, 2] == "N"
                   If !this.Write(Space(24) + '<xsd:element name="' + Lower(aCampos[nPos, 1]) + '">')
                      Return(.F.)
                   Endif

                   If !this.Write(Space(27) + '<xsd:simpleType>')
                      Return(.F.)
                   Endif

                   If !this.Write(Space(30) + '<xsd:restriction base="xsd:decimal">')
                      Return(.F.)
                   Endif

                   If !this.Write(Space(33) + '<xsd:totalDigits value="' + LTrim(Str(aCampos[nPos, 3], 3, 0)) +'"/>')
                      Return(.F.)
                   Endif

                   If !this.Write(Space(33) + '<xsd:fractionDigits value="' + LTrim(Str(aCampos[nPos, 4], 3, 0)) +'"/>')
                      Return(.F.)
                   Endif

                   If !this.Write(Space(30) + '</xsd:restriction>')
                      Return(.F.)
                   Endif
 
                   If !this.Write(Space(27) + '</xsd:simpleType>')
                      Return(.F.)
                   Endif

                   If !this.Write(Space(24) + '</xsd:element>')
                      Return(.F.)
                   Endif


              Case aCampos[nPos, 2] == "F"
                   If !this.Write(Space(24) + '<xsd:element name="' + Lower(aCampos[nPos, 1]) + '">')
                      Return(.F.)
                   Endif

                   If !this.Write(Space(27) + '<xsd:simpleType>')
                      Return(.F.)
                   Endif

                   If !this.Write(Space(30) + '<xsd:restriction base="xsd:decimal">')
                      Return(.F.)
                   Endif

                   If !this.Write(Space(33) + '<xsd:totalDigits value="' + LTrim(Str(aCampos[nPos, 3], 3, 0)) +'"/>')
                      Return(.F.)
                   Endif

                   If !this.Write(Space(33) + '<xsd:fractionDigits value="' + LTrim(Str(aCampos[nPos, 4], 3, 0)) +'"/>')
                      Return(.F.)
                   Endif

                   If !this.Write(Space(30) + '</xsd:restriction>')
                      Return(.F.)
                   Endif
 
                   If !this.Write(Space(27) + '</xsd:simpleType>')
                      Return(.F.)
                   Endif

                   If !this.Write(Space(24) + '</xsd:element>')
                      Return(.F.)
                   Endif


              Case aCampos[nPos, 2] == "I"
                   If !this.Write(Space(24) + '<xsd:element name="' + Lower(aCampos[nPos, 1]) + '" type="xsd:int"/>')
                      Return(.F.)
                   Endif


              Case aCampos[nPos, 2] == "B"
                   If !this.Write(Space(24) + '<xsd:element name="' + Lower(aCampos[nPos, 1]) + '" type="xsd:double"/>')
                      Return(.F.)
                   Endif


              Case aCampos[nPos, 2] == "Y"
                   If !this.Write(Space(24) + '<xsd:element name="' + Lower(aCampos[nPos, 1]) + '">')
                      Return(.F.)
                   Endif

                   If !this.Write(Space(27) + '<xsd:simpleType>')
                      Return(.F.)
                   Endif

                   If !this.Write(Space(30) + '<xsd:restriction base="xsd:decimal">')
                      Return(.F.)
                   Endif

                   If !this.Write(Space(33) + '<xsd:totalDigits value="19"/>')
                      Return(.F.)
                   Endif

                   If !this.Write(Space(33) + '<xsd:fractionDigits value="4"/>')
                      Return(.F.)
                   Endif

                   If !this.Write(Space(30) + '</xsd:restriction>')
                      Return(.F.)
                   Endif
 
                   If !this.Write(Space(27) + '</xsd:simpleType>')
                      Return(.F.)
                   Endif

                   If !this.Write(Space(24) + '</xsd:element>')
                      Return(.F.)
                   Endif


              Case aCampos[nPos, 2] == "T"
                   If !this.Write(Space(24) + '<xsd:element name="' + Lower(aCampos[nPos, 1]) + '" type="xsd:dateTime"/>')
                      Return(.F.)
                   Endif

           EndCase
        Endif

    Next


    If Empty(cFile)
       cRes = cRes + Space(21) + '</xsd:sequence>' + NEW_LINE
       cRes = cRes + Space(18) + '</xsd:complexType>' + NEW_LINE
       cRes = cRes + Space(15) + '</xsd:element>' + NEW_LINE
    else
       If !this.Write(Space(21) + '</xsd:sequence>')
          Return(.F.)
       Endif

       If !this.Write(Space(18) + '</xsd:complexType>')
          Return(.F.)
       Endif

       If !this.Write(Space(15) + '</xsd:element>')
          Return(.F.)
       Endif
    Endif
Next

If Empty(cFile)
   cRes = cRes + Space(12) + '</xsd:choice>' + NEW_LINE
   cRes = cRes + Space(9) + '</xsd:complexType>' + NEW_LINE
   cRes = cRes + Space(6) + '</xsd:element>' + NEW_LINE
   cRes = cRes + Space(3) + '</xsd:schema>' + NEW_LINE
else
   If !this.Write(Space(12) + '</xsd:choice>')
      Return(.F.)
   Endif

   If !this.Write(Space(9) + '</xsd:complexType>')
      Return(.F.)
   Endif

   If !this.Write(Space(6) + '</xsd:element>')
      Return(.F.)
   Endif

   If !this.Write(Space(3) + '</xsd:schema>')
      Return(.F.)
   Endif
Endif

Endif



* Generar XML
For nCont = 2 to ALen(this.Tables)
    Select(this.Tables[nCont])

    Dimension aCampos[1]
    nFields = AFields(aCampos)

    Scan

       If Empty(cFile)
          cRes = cRes + Space(3) + "<" + this.Tables[nCont] + ">" + NEW_LINE
       else
          If !this.Write(Space(3) + "<" + this.Tables[nCont] + ">")
             Return(.F.)
          Endif
       Endif


       For nPos = 1 to nFields

           If aCampos[nPos, 2] == "G"
              Loop
           Endif

           xValue = Evaluate(aCampos[nPos, 1])

           If (Empty(xValue) .OR. IsNull(xValue)) .AND. (!InList(aCampos[nPos, 2], "L", "I", "B", "Y"))
              cItem = Space(6) + "<" + Lower(aCampos[nPos, 1]) + "/>"
           else
              cItem = Space(6) + "<" + Lower(aCampos[nPos, 1]) + ">"

              Do Case
                 Case aCampos[nPos, 2] == "C"
                      cItem = cItem + this.CleanXML(RTrim(xValue))


                 Case aCampos[nPos, 2] == "D"
                      cItem = cItem + PadL(Year(xValue), 4, "0") + "-"
                      cItem = cItem + PadL(Month(xValue), 2, "0") + "-"
                      cItem = cItem + PadL(Day(xValue), 2, "0")


                 Case aCampos[nPos, 2] == "L"
                      If xValue
                         cItem = cItem + "true"
                      else
                         cItem = cItem + "false"
                      Endif


                 Case aCampos[nPos, 2] == "M"
                      cItem = cItem + "<![CDATA["
                      cItem = cItem + RTrim(xValue)
                      cItem = cItem + "]]>"


                 Case aCampos[nPos, 2] == "N"
                      Set Point To "."
                      Set Separator To ","
                      cItem = cItem + LTrim(Str(xValue, aCampos[nPos, 3], aCampos[nPos, 4]))
                      Set Point To &cAntPoint.
                      Set Separator To &cAntSeparator.


                 Case aCampos[nPos, 2] == "F"
                      Set Point To "."
                      Set Separator To ","
                      cItem = cItem + LTrim(Str(xValue, aCampos[nPos, 3], aCampos[nPos, 4]))
                      Set Point To &cAntPoint.
                      Set Separator To &cAntSeparator.


                 Case aCampos[nPos, 2] == "I"
                      cItem = cItem + LTrim(Str(xValue, 18, 0))


                 Case aCampos[nPos, 2] == "B"
                      Set Point To "."
                      Set Separator To ","
                      cItem = cItem + LTrim(Str(xValue, 18, aCampos[nPos, 4]))
                      Set Point To &cAntPoint.
                      Set Separator To &cAntSeparator.


                 Case aCampos[nPos, 2] == "Y"
                      Set Point To "."
                      Set Separator To ","
                      cItem = cItem + LTrim(Str(xValue, 18, 4))
                      Set Point To &cAntPoint.
                      Set Separator To &cAntSeparator.


                 Case aCampos[nPos, 2] == "T"
                      cItem = cItem + PadL(Year(xValue), 4, "0") + "-"
                      cItem = cItem + PadL(Month(xValue), 2, "0") + "-"
                      cItem = cItem + PadL(Day(xValue), 2, "0") + "T"
                      cItem = cItem + PadL(Hour(xValue), 2, "0") + ":"
                      cItem = cItem + PadL(Minute(xValue), 2, "0") + ":"
                      cItem = cItem + PadL(Sec(xValue), 2, "0")

              EndCase

              cItem = cItem + "</" + Lower(aCampos[nPos, 1]) + ">"
           Endif

           If Empty(cFile)
              cRes = cRes + cItem + NEW_LINE
           else
              If !this.Write(cItem)
                 Return(.F.)
              Endif
           Endif

       Next


       If Empty(cFile)
          cRes = cRes + Space(3) + "</" + this.Tables[nCont] + ">" + NEW_LINE
       else
          If !this.Write(Space(3) + "</" + this.Tables[nCont] + ">")
             Return(.F.)
          Endif
       Endif

    EndScan
Next


If Empty(cFile)
   cRes = cRes + "</VFPData>" + NEW_LINE
else
   If !this.Write("</VFPData>")
      Return(.F.)
   Endif

   FClose(this.Handle)
Endif


If Empty(cFile)
   Return(cRes)
else
   Return(.T.)
Endif

EndFunc



***
*  Function Write(cString)
***
Hidden Function Write(cString)

FWrite(this.Handle, cString + NEW_LINE)

EndFunc



***
*  Function CleanXML(cString)
***
Hidden Function CleanXML(cString)

cString = StrTran(cString, "&", "&amp;")
cString = StrTran(cString, "<", "&lt;")
cString = StrTran(cString, ">", "&gt;")

Return(cString)



***
*  Function WriteStartDocument()
***
Function WriteStartDocument()

this.TempXML = '<?xml version = "1.0" encoding="Windows-1252" standalone="yes"?>' + NEW_LINE
this.TempXML = this.TempXML + "<VFPData>" + NEW_LINE
EndFunc



***
*  Function WriteEndDocument()
***
Function WriteEndDocument()

this.TempXML = this.TempXML + "</VFPData>" + NEW_LINE

EndFunc



***
*  Function WriteStartElement(cName)
***
Function WriteStartElement(cName)

this.TempItemName = cName
this.TempXML = this.TempXML + Space(3) + "<" + cName + ">" + NEW_LINE

EndFunc



***
*  Function WriteEndElement()
***
Function WriteEndElement()

this.TempXML = this.TempXML + Space(3) + "</" + this.TempItemName + ">" + NEW_LINE

EndFunc



***
*  Function WriteElementString(cName, cValue)
***
Function WriteElementString(cName, cValue)

If Len(cValue) == 0
   this.TempXML = this.TempXML + Space(6) + "<" + cName + "/>" + NEW_LINE
else
   this.TempXML = this.TempXML + Space(6) + "<" + cName + ">" + cValue + ;
                                 "</" + cName + ">" +  + NEW_LINE
Endif

EndFunc



***
*  Function GetXML()
***
Function GetXML()

Return(this.TempXML)



***
*  XMLToCursor(cXML, cCursorList)
***
Function XMLToCursor(cXML, cCursorList)
Local nCont, oXML, oRoot, oSchema, lHasSchema, aCursors, aNames
Local oCursors, oCursor, nNode, nAntWA
Local nCont1, nCont2, nCont3, nCont4
Local oFields, oField, oFieldDetail, oFieldDetailItem
Local cType, nLen, nDecimals
Local cName, nPos, cLetter
Local cCommand, lFound, cText
Local cAntNear, cAntDate
Local nProcessedBatch, nProcessedTotal, nProcessedCount


* Inicializar propiedades
If Type("cCursorList") <> "C"
   cCursorList = ""
Endif


* Guardar entorno
nAntWA = Select(0)

* Inicializar parser
oXML = CreateObject(this.Parser)

oXML.LoadXML(cXML)
oXML.PreserveWhiteSpace = .T.


* Buscar raiz
oRoot = oXML.DocumentElement

If IsNull(oRoot)
   Select(nAntWA)
   Return(.F.)
Endif


* Leer esquema
Dimension aCursors[1]
aCursors[1] = "xsd:"

Create Cursor Schemas (Cursor C(50), Columna C(50), ;
                       Type C(1), Len N(10, 0), Decimals N(10, 0))
Index on Cursor + Columna Tag Columna

lHasSchema = .F.

For nCont = 0 to oRoot.ChildNodes.Length - 1
    If Upper(Left(oRoot.ChildNodes(nCont).TagName, 10)) == Upper("xsd:schema")
       lHasSchema = .T.

       oSchema = oRoot.ChildNodes(nCont)
       For nCont1 = 0 to oSchema.ChildNodes.Length - 1
           If Upper(Left(oSchema.ChildNodes(nCont1).TagName, 11)) == Upper("xsd:element")

              If oSchema.ChildNodes(nCont1).GetAttribute("name") == oRoot.TagName

                 oCursors = oSchema.ChildNodes(nCont1).ChildNodes(0).ChildNodes(0)

                 For nCont2 = 0 to oCursors.ChildNodes.Length - 1
                     oCursor = oCursors.ChildNodes(nCont2)
                     Dimension aCursors[ALen(aCursors) + 1]
                     aCursors[ALen(aCursors)] = oCursor.GetAttribute("name")

                     oFields = oCursor.ChildNodes(0).ChildNodes(0)

                     For nCont3 = 0 to oFields.ChildNodes.Length - 1

                         oField = oFields.ChildNodes(nCont3)

                         cType = "C"
                         nLen = 0
                         nDecimals = 0

                         Do Case
                            Case Upper(oField.GetAttribute("type")) == Upper("xsd:date")
                                 cType = "D"
                            Case Upper(oField.GetAttribute("type")) == Upper("xsd:dateTime")
                                 cType = "T"
                            Case Upper(oField.GetAttribute("type")) == Upper("xsd:boolean")
                                 cType = "L"
                            Case Upper(oField.GetAttribute("type")) == Upper("xsd:int")
                                 cType = "I"
                            Case Upper(oField.GetAttribute("type")) == Upper("xsd:double")
                                 cType = "B"

                            OtherWise
                                 oFieldDetail = oField.ChildNodes(0).ChildNodes(0)
                                 Do Case
                                    Case Upper(oFieldDetail.GetAttribute("base")) == Upper("xsd:decimal")

                                         cType = "N"
                                         For nCont4 = 0 to oFieldDetail.ChildNodes.Length - 1
                                             oFieldDetailItem = oFieldDetail.ChildNodes(nCont4)

                                             Do Case
                                                Case Upper(oFieldDetailItem.TagName) == Upper("xsd:totalDigits")
                                                     nLen = Val(oFieldDetailItem.GetAttribute("value"))
                                                     If IsNull(nLen)
                                                        nLen = 0
                                                     Endif

                                                Case Upper(oFieldDetailItem.TagName) == Upper("xsd:fractionDigits")
                                                     nDecimals = Val(oFieldDetailItem.GetAttribute("value"))
                                                     If IsNull(nDecimals)
                                                        nDecimals = 0
                                                     Endif

                                             EndCase

                                         Next


                                    Case Upper(oFieldDetail.GetAttribute("base")) == Upper("xsd:string")
                                         cType = "C"
                                         For nCont4 = 0 to oFieldDetail.ChildNodes.Length - 1
                                             oFieldDetailItem = oFieldDetail.ChildNodes(nCont4)

                                             If Upper(oFieldDetailItem.TagName) == Upper("xsd:maxLength")
                                                nLen = Val(oFieldDetailItem.GetAttribute("value"))
 
                                                If IsNull(nLen)
                                                   nLen = 0
                                                 Endif
 
                                                 If nLen = 2147483647
                                                    cType = "M"
                                                    nLen = 0
                                                 Endif
                                             Endif

                                         Next

                                 EndCase
                         EndCase


                         Select "Schemas"
                         Set Order to Tag Columna
                         If !Seek(PadR(aCursors[ALen(aCursors)], 50) + PadR(oField.GetAttribute("name"), 50))
                            Append Blank
                            Replace Schemas.Cursor with aCursors[ALen(aCursors)]
                            Replace Schemas.Columna with oField.GetAttribute("name")
                         Endif

                         Replace Schemas.Type with cType
                         Replace Schemas.Len with nLen
                         Replace Schemas.Decimals with nDecimals
                     Next

                 Next

              Endif
           Endif
       Next

       Exit
    Endif
Next


* Determinar los cursores a leer
If !lHasSchema
   For nCont = 0 to oRoot.ChildNodes.Length - 1
       If Upper(Left(oRoot.ChildNodes(nCont).TagName, 10)) <> Upper("xsd:schema")
          If AScan(aCursors, oRoot.ChildNodes(nCont).TagName) == 0
             Dimension aCursors[ALen(aCursors) + 1]
             aCursors[ALen(aCursors)] = oRoot.ChildNodes(nCont).TagName
          Endif
       Endif
   Next
Endif


* Generar Nombres
If Empty(cCursorList)
   Dimension aNames[ALen(aCursors)]
   For nCont = 2 to ALen(aCursors)
       aNames[nCont] = aCursors[nCont]
   Next
else
   Dimension aNames[1]
   Do While Len(cCursorList) <> 0
      Dimension aNames[ALen(aNames) + 1]

      If At(",", cCursorList) == 0
         aNames[ALen(aNames)] = AllTrim(cCursorList)
         cCursorList = ""
      else
         aNames[ALen(aNames)] = AllTrim(Left(cCursorList, At(",", cCursorList) - 1))
         cCursorList = SubStr(cCursorList, At(",", cCursorList) + 1)
      Endif
   Enddo

   If ALen(aNames) < ALen(aCursors)
      Dimension aCursors[ALen(aNames)]
   Endif
Endif


* Validar Nombres
For nCont = 2 to ALen(aNames)

    cName = ""

    For nPos = 1 to Len(aNames[nCont])
        cLetter = SubStr(aNames[nCont], nPos, 1)

        Do Case
           Case (cLetter >= "0") .AND. (cLetter <= "9")
                * Letra Valida

           Case (cLetter >= "A") .AND. (cLetter <= "Z")
                * Letra Valida

           Case (cLetter >= "a") .AND. (cLetter <= "z")
                * Letra Valida

           Case cLetter == "_"
                * Letra Valida

           OtherWise
                cLetter = "_"

        EndCase

        cName = cName + cLetter
    Next

    aNames[nCont] = cName

Next


* Generar esquema dinamicamente
If !lHasSchema
   For nCont = 2 to ALen(aCursors)
       oCursors = oRoot.SelectNodes(aCursors[nCont])
       For Each oCursor in oCursors

           For nNode = 0 to oCursor.ChildNodes.Length - 1

               Select "Schemas"
               Set Order to Tag Columna
               If !Seek(PadR(aCursors[nCont], 50) + PadR(oCursor.ChildNodes(nNode).TagName, 50))
                  Append Blank
                  Replace Schemas.Cursor with aCursors[nCont]
                  Replace Schemas.Columna with oCursor.ChildNodes(nNode).TagName
                  Replace Schemas.Type with "C"
               Endif

               If Len(oCursor.ChildNodes(nNode).Text) > Schemas.Len
                  Replace Schemas.Len with Len(oCursor.ChildNodes(nNode).Text)

                  If (Schemas.Type == "C") .AND. (Schemas.Len > 254)
                     Replace Schemas.Type with "M"
                  Endif
               Endif

           Next

       Next
   Next
Endif

Replace Schemas.Len with 10 For Schemas.Len <= 0


* Crear Cursores
For nCont = 2 to ALen(aNames)

    If ALen(aCursors) < nCont
       Exit
    Endif
 
    cCommand = "Create Cursor " + aNames[nCont] + "("

    lFound = .F.

    cAntNear = Set("Near")
    Set Near On
    Select "Schemas"
    Seek(PadR(aCursors[nCont], 50) + Space(50))
    Set Near &cAntNear.

    Do While (!Eof()) .AND. (Schemas.Cursor = aCursors[nCont])

       cCommand = cCommand + AllTrim(Schemas.Columna) + " " + ;
                             Schemas.Type + "(" + ;
                             AllTrim(Str(Schemas.Len, 10, 0))

       If Schemas.Decimals <> 0
          cCommand = cCommand + "," + AllTrim(Str(Schemas.Decimals, 10, 0))
       Endif

       cCommand = cCommand + "),"

       lFound = .T.
       Skip
    Enddo

    cCommand = Left(cCommand, Len(cCommand) - 1) + ")"

    If !lFound
       Select(nAntWA)
       Return(.F.)
    else
       &cCommand.
    Endif

Next


* Determinar la cantidad total de registros a leer
nProcessedTotal = 0
nProcessedCount = 0

For nCont = 2 to ALen(aCursors)
    oCursors = oRoot.SelectNodes(aCursors[nCont])
    nProcessedTotal = nProcessedTotal + oCursors.length
Next


* Llenar Cursores
For nCont = 2 to ALen(aCursors)

    Select (aNames[nCont])
    oCursors = oRoot.SelectNodes(aCursors[nCont])

    nProcessedBatch = 0
    For Each oCursor in oCursors

        nProcessedCount = nProcessedCount + 1

        Append Blank
        For nNode = 0 to oCursor.ChildNodes.Length - 1
            If lHasSchema

               If Seek(PadR(aCursors[nCont], 50) + PadR(oCursor.ChildNodes(nNode).TagName, 50), "Schemas")

                  Do Case
                     Case Schemas.Type == "C"
                          Replace (oCursor.ChildNodes(nNode).TagName) with oCursor.ChildNodes(nNode).Text

                     Case Schemas.Type == "D"
                          cText = Right(oCursor.ChildNodes(nNode).Text, 2) + "/"
                          cText = cText + SubStr(oCursor.ChildNodes(nNode).Text, 6, 2) + "/"
                          cText = cText + Left(oCursor.ChildNodes(nNode).Text, 4)

                          cAntDate = Set("Date")
                          Set Date to French
                          Replace (oCursor.ChildNodes(nNode).TagName) with CtoD(cText)
                          Set Date to (cAntDate)

                     Case Schemas.Type == "L"
                          Replace (oCursor.ChildNodes(nNode).TagName) with IIF(Upper(oCursor.ChildNodes(nNode).Text) == "TRUE", .T., .F.)

                     Case Schemas.Type == "M"
                          Replace (oCursor.ChildNodes(nNode).TagName) with oCursor.ChildNodes(nNode).Text

                     Case Schemas.Type == "N"
                          Replace (oCursor.ChildNodes(nNode).TagName) with Val(oCursor.ChildNodes(nNode).Text)

                     Case Schemas.Type == "F"
                          Replace (oCursor.ChildNodes(nNode).TagName) with Val(oCursor.ChildNodes(nNode).Text)

                     Case Schemas.Type == "I"
                          Replace (oCursor.ChildNodes(nNode).TagName) with Val(oCursor.ChildNodes(nNode).Text)

                     Case Schemas.Type == "B"
                          Replace (oCursor.ChildNodes(nNode).TagName) with Val(oCursor.ChildNodes(nNode).Text)

                     Case Schemas.Type == "Y"
                          Replace (oCursor.ChildNodes(nNode).TagName) with Val(oCursor.ChildNodes(nNode).Text)

                     Case Schemas.Type == "T"
                          cText = SubStr(oCursor.ChildNodes(nNode).Text, 9, 2) + "/"
                          cText = cText + SubStr(oCursor.ChildNodes(nNode).Text, 6, 2) + "/"
                          cText = cText + Left(oCursor.ChildNodes(nNode).Text, 4) + " "
                          cText = cText + Right(oCursor.ChildNodes(nNode).Text, 8)

                          cAntDate = Set("Date")
                          Set Date to French
                          Replace (oCursor.ChildNodes(nNode).TagName) with CtoT(cText)
                          Set Date to (cAntDate)

                  EndCase
    
               Endif

            else
               Replace (oCursor.ChildNodes(nNode).TagName) with oCursor.ChildNodes(nNode).Text
            Endif
        Next

        nProcessedBatch = nProcessedBatch + 1
        If nProcessedBatch >= this.Interval
           If Type("this.Status") == "O"
              this.Status.UpdateStatus(2, nProcessedTotal, nProcessedCount)
              nProcessedBatch = 0
           Endif
        Endif

    Next
Next


* Actualizar el puntero de los cursores
Use in Schemas

For nCont = 2 to ALen(aNames)
    If Used(aNames[nCont])
       Go Top in (aNames[nCont])
    Endif
Next


* Restablecer entorno
Select(nAntWA)

Return(.T.)



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



***
*  Function ValidCursor(cCursor, cFieldList)
***
Function ValidCursor(cCursor, cFieldList)
Local aCampos, nCampos, cField, nCont, lFound

If !Used(cCursor)
   Return(.F.)
Endif

Dimension aCampos[1]
nCampos = AFields(aCampos, cCursor)

Do While !Empty(cFieldList)

   If At(",", cFieldList) == 0
      cField = Upper(cFieldList)
      cFieldList = ""
   else
      cField = Upper(Left(cFieldList, At(",", cFieldList) - 1))
      cFieldList = SubStr(cFieldList, At(",", cFieldList) + 1)
   Endif

   lFound = .F.
   For nCont = 1 to nCampos
       If Upper(aCampos[nCont, 1]) == cField
          lFound = .T.
          Exit
       Endif
   Next

   If !lFound
      Return(.F.)
   Endif

Enddo

Return(.T.)


EndDefine
