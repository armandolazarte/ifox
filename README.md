# iFox
A set of tools for Microsoft Visual FoxPro implementing the most common internet protocols.


This project started as a freeware library in www.ifox.com.ar around 1999 to demonstrate what can be done in Microsoft Visual FoxPro.
As years passes by it evolved as a complete solution, covering a broad range of protocols.

As Pablo Pioli, the original developer, moved to other horizons the source code went public with a LGPL license.

It's important to remark that the project will not receive significant additions but pull requests will be accepted. Includes a (mostly) updated documentation.



Includes the following components:


**HTTP**

Get any URL source, simulate forms or do a file upload. 


**POP3**

Read the messages of any POP3 mail account with this component. You can also do selective deletes, download only the headings and handle file attachments. 


**SMTP**

This component allows you to send E-Mail messages from you applications without the need to use a mail program. Attachments support is also included. 


**FTP**

Build a full featured FTP client with this component that allows you to download and upload files, create folder, list their content, rename and delete files and much more. This components uses the well known WinInet library from Microsoft. 


**DirectFTP**

This components can do the same tasks than iFox.FTP but was built from the ground by the iFox team, so it does not suffer any of the WinInet limitations. Using iFox.DirectFTP is possible to continue interrupted downloads, move files and add information to an existing file. 


**Resume&Go**

This component allows to start a download and continue it in another moment like the most popular download managers. 


**Sockets**

This component wraps the Winsock 2 functionality in a simple and poweful package. Stablish multiple connections simultaneosly and use multiple ports. Build a server with just a few lines of code.


**Connect**

Stablish and close Internet connections with this component. You can also monitor any connection statistics. 


**XML**

Using iFox.XML is possible to convert one or multiple cursors to the XML format or the inverse. You can also build an XML file dinamically. 


**NTP**

Communicates with time servers. 
