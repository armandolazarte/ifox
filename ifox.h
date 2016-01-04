
#define FINAL_VERSION "SI"


* Definiciones varias
#define NEW_LINE  Chr(13) + Chr(10)


* Constantes para MessageBox
#define MB_YESNO                 4
#define MB_ICONQUESTION         32
#define IDYES                    6


* Parametro no especificado
#define PAR_NOTSPECIFIED   -1


* Registry roots
#define HKEY_CLASSES_ROOT           -2147483648  && BITSET(0,31)
#define HKEY_CURRENT_USER           -2147483647  && BITSET(0,31) + 1
#define HKEY_LOCAL_MACHINE          -2147483646  && BITSET(0,31) + 2
#define HKEY_USERS                  -2147483645  && BITSET(0,31) + 3
#define HKEY_DYN_DATA               -2147483642  && BITSET(0,31) + 6





* Lenguajes
#define LANG_ESPANOL        1
#define LANG_ENGLISH        2



* Ping Return Codes
#define INADDR_NONE                -1
#define IP_BUF_TOO_SMALL           (11000 + 1)
#define IP_DEST_NET_UNREACHABLE    (11000 + 2)
#define IP_DEST_HOST_UNREACHABLE   (11000 + 3)
#define IP_DEST_PROT_UNREACHABLE   (11000 + 4)
#define IP_DEST_PORT_UNREACHABLE   (11000 + 5)
#define IP_NO_RESOURCES            (11000 + 6)
#define IP_BAD_OPTION              (11000 + 7)
#define IP_HW_ERROR                (11000 + 8)
#define IP_PACKET_TOO_BIG          (11000 + 9)
#define IP_REQ_TIMED_OUT           (11000 + 10)
#define IP_BAD_REQ                 (11000 + 11)
#define IP_BAD_ROUTE               (11000 + 12)
#define IP_TTL_EXPIRED_TRANSIT     (11000 + 13)
#define IP_TTL_EXPIRED_REASSEM     (11000 + 14)
#define IP_PARAM_PROBLEM           (11000 + 15)
#define IP_SOURCE_QUENCH           (11000 + 16)
#define IP_OPTION_TOO_BIG          (11000 + 17)
#define IP_BAD_DESTINATION         (11000 + 18)
#define IP_ADDR_DELETED            (11000 + 19)
#define IP_SPEC_MTU_CHANGE         (11000 + 20)
#define IP_MTU_CHANGE              (11000 + 21)
#define IP_UNLOAD                  (11000 + 22)
#define IP_ADDR_ADDED              (11000 + 23)
#define IP_GENERAL_FAILURE         (11000 + 50)
#define IP_PENDING                 (11000 + 255)
#define PING_TIMEOUT               500


* SMTP Priorities
#define SMTP_PRIORITY_HIGH         1
#define SMTP_PRIORITY_NORMAL       3
#define SMTP_PRIORITY_LOW          5


* Body Types
#define SMTP_BODY_STRING           1
#define SMTP_BODY_FILE             2


* Download: Type
#define DOWNLOAD_STRING            1
#define DOWNLOAD_FILE              2


* iFox.Resume&Go: TransferType property
#define TYPE_FILE                  1
#define TYPE_MEMORY                2


* XSD Types
#define XSD_NONE       0
#define XSD_INLINE     1
#define XSD_EXTERNAL   2
