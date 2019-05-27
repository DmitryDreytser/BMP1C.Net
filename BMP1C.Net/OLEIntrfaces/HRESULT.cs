using System;
using System.Linq;

namespace BMP1C.Net
{
    public enum HRESULT : uint
    {
        S_OK = 0x00000000,
        S_FALSE = 0x00000001,
        STG_E_INVALIDFUNCTION = 0x80030001,
        STG_E_FILENOTFOUND = 0x80030002,
        STG_E_PATHNOTFOUND = 0x80030003,
        STG_E_TOOMANYOPENFILES = 0x80030004,
        STG_E_ACCESSDENIED = 0x80030005,
        STG_E_INVALIDHANDLE = 0x80030006,
        STG_E_INSUFFICIENTMEMORY = 0x80030008,
        STG_E_INVALIDPOINTER = 0x80030009,
        STG_E_NOMOREFILES = 0x80030012,
        STG_E_DISKISWRITEPROTECTED = 0x80030013,
        STG_E_SEEKERROR = 0x80030019,
        STG_E_WRITEFAULT = 0x8003001D,
        STG_E_READFAULT = 0x8003001E,
        STG_E_SHAREVIOLATION = 0x80030020,
        STG_E_LOCKVIOLATION = 0x80030021,
        STG_E_FILEALREADYEXISTS = 0x80030050,
        STG_E_INVALIDPARAMETER = 0x80030057,
        STG_E_MEDIUMFULL = 0x80030070,
        STG_E_PROPSETMISMATCHED = 0x800300F0,
        STG_E_ABNORMALAPIEXIT = 0x800300FA,
        STG_E_INVALIDHEADER = 0x800300FB,
        STG_E_INVALIDNAME = 0x800300FC,
        STG_E_UNKNOWN = 0x800300FD,
        STG_E_UNIMPLEMENTEDFUNCTION = 0x800300FE,
        STG_E_INVALIDFLAG = 0x800300FF,
        STG_E_INUSE = 0x80030100,
        STG_E_NOTCURRENT = 0x80030101,
        STG_E_REVERTED = 0x80030102,
        STG_E_CANTSAVE = 0x80030103,
        STG_E_OLDFORMAT = 0x80030104,
        STG_E_OLDDLL = 0x80030105,
        STG_E_SHAREREQUIRED = 0x80030106,
        STG_E_NOTFILEBASEDSTORAGE = 0x80030107,
        STG_E_EXTANTMARSHALLINGS = 0x80030108,
        STG_E_DOCFILECORRUPT = 0x80030109,
        STG_E_BADBASEADDRESS = 0x80030110,
        STG_E_DOCFILETOOLARGE = 0x80030111,
        STG_E_NOTSIMPLEFORMAT = 0x80030112,
        STG_E_INCOMPLETE = 0x80030201,
        STG_E_TERMINATED = 0x80030202,
        STG_S_CONVERTED = 0x00030200,
        STG_S_BLOCK = 0x00030201,
        STG_S_RETRYNOW = 0x00030202,
        STG_S_MONITORING = 0x00030203,
        STG_S_MULTIPLEOPENS = 0x00030204,
        STG_S_CONSOLIDATIONFAILED = 0x00030205,
        STG_S_CANNOTCONSOLIDATE = 0x00030206,
        STG_E_STATUS_COPY_PROTECTION_FAILURE = 0x80030305,
        STG_E_CSS_AUTHENTICATION_FAILURE = 0x80030306,
        STG_E_CSS_KEY_NOT_PRESENT = 0x80030307,
        STG_E_CSS_KEY_NOT_ESTABLISHED = 0x80030308,
        STG_E_CSS_SCRAMBLED_SECTOR = 0x80030309,
        STG_E_CSS_REGION_MISMATCH = 0x8003030A,
        STG_E_RESETS_EXHAUSTED = 0x8003030B,
        RPC_E_CALL_REJECTED = 0x80010001,
        RPC_E_CALL_CANCELED = 0x80010002,
        RPC_E_CANTPOST_INSENDCALL = 0x80010003,
        RPC_E_CANTCALLOUT_INASYNCCALL = 0x80010004,
        RPC_E_CANTCALLOUT_INEXTERNALCALL = 0x80010005,
        RPC_E_CONNECTION_TERMINATED = 0x80010006,
        RPC_E_SERVER_DIED = 0x80010007,
        RPC_E_CLIENT_DIED = 0x80010008,
        RPC_E_INVALID_DATAPACKET = 0x80010009,
        RPC_E_CANTTRANSMIT_CALL = 0x8001000A,
        RPC_E_CLIENT_CANTMARSHAL_DATA = 0x8001000B,
        RPC_E_CLIENT_CANTUNMARSHAL_DATA = 0x8001000C,
        RPC_E_SERVER_CANTMARSHAL_DATA = 0x8001000D,
        RPC_E_SERVER_CANTUNMARSHAL_DATA = 0x8001000E,
        RPC_E_INVALID_DATA = 0x8001000F,
        RPC_E_INVALID_PARAMETER = 0x80010010,
        RPC_E_CANTCALLOUT_AGAIN = 0x80010011,
        RPC_E_SERVER_DIED_DNE = 0x80010012,
        RPC_E_SYS_CALL_FAILED = 0x80010100,
        RPC_E_OUT_OF_RESOURCES = 0x80010101,
        RPC_E_ATTEMPTED_MULTITHREAD = 0x80010102,
        RPC_E_NOT_REGISTERED = 0x80010103,
        RPC_E_FAULT = 0x80010104,
        RPC_E_SERVERFAULT = 0x80010105,
        RPC_E_CHANGED_MODE = 0x80010106,
        RPC_E_INVALIDMETHOD = 0x80010107,
        RPC_E_DISCONNECTED = 0x80010108,
        RPC_E_RETRY = 0x80010109,
        RPC_E_SERVERCALL_RETRYLATER = 0x8001010A,
        RPC_E_SERVERCALL_REJECTED = 0x8001010B,
        RPC_E_INVALID_CALLDATA = 0x8001010C,
        RPC_E_CANTCALLOUT_ININPUTSYNCCALL = 0x8001010D,
        RPC_E_WRONG_THREAD = 0x8001010E,
        RPC_E_THREAD_NOT_INIT = 0x8001010F,
        RPC_E_VERSION_MISMATCH = 0x80010110,
        RPC_E_INVALID_HEADER = 0x80010111,
        RPC_E_INVALID_EXTENSION = 0x80010112,
        RPC_E_INVALID_IPID = 0x80010113,
        RPC_E_INVALID_OBJECT = 0x80010114,
        RPC_S_CALLPENDING = 0x80010115,
        RPC_S_WAITONTIMER = 0x80010116,
        RPC_E_CALL_COMPLETE = 0x80010117,
        RPC_E_UNSECURE_CALL = 0x80010118,
        RPC_E_TOO_LATE = 0x80010119,
        RPC_E_NO_GOOD_SECURITY_PACKAGES = 0x8001011A,
        RPC_E_ACCESS_DENIED = 0x8001011B,
        RPC_E_REMOTE_DISABLED = 0x8001011C,
        RPC_E_INVALID_OBJREF = 0x8001011D,
        RPC_E_NO_CONTEXT = 0x8001011E,
        RPC_E_TIMEOUT = 0x8001011F,
        RPC_E_NO_SYNC = 0x80010120,
        RPC_E_FULLSIC_REQUIRED = 0x80010121,
        RPC_E_INVALID_STD_NAME = 0x80010122,
        CO_E_FAILEDTOIMPERSONATE = 0x80010123,
        CO_E_FAILEDTOGETSECCTX = 0x80010124,
        CO_E_FAILEDTOOPENTHREADTOKEN = 0x80010125,
        CO_E_FAILEDTOGETTOKENINFO = 0x80010126,
        CO_E_TRUSTEEDOESNTMATCHCLIENT = 0x80010127,
        CO_E_FAILEDTOQUERYCLIENTBLANKET = 0x80010128,
        CO_E_FAILEDTOSETDACL = 0x80010129,
        CO_E_ACCESSCHECKFAILED = 0x8001012A,
        CO_E_NETACCESSAPIFAILED = 0x8001012B,
        CO_E_WRONGTRUSTEENAMESYNTAX = 0x8001012C,
        CO_E_INVALIDSID = 0x8001012D,
        CO_E_CONVERSIONFAILED = 0x8001012E,
        CO_E_NOMATCHINGSIDFOUND = 0x8001012F,
        CO_E_LOOKUPACCSIDFAILED = 0x80010130,
        CO_E_NOMATCHINGNAMEFOUND = 0x80010131,
        CO_E_LOOKUPACCNAMEFAILED = 0x80010132,
        CO_E_SETSERLHNDLFAILED = 0x80010133,
        CO_E_FAILEDTOGETWINDIR = 0x80010134,
        CO_E_PATHTOOLONG = 0x80010135,
        CO_E_FAILEDTOGENUUID = 0x80010136,
        CO_E_FAILEDTOCREATEFILE = 0x80010137,
        CO_E_FAILEDTOCLOSEHANDLE = 0x80010138,
        CO_E_EXCEEDSYSACLLIMIT = 0x80010139,
        CO_E_ACESINWRONGORDER = 0x8001013A,
        CO_E_INCOMPATIBLESTREAMVERSION = 0x8001013B,
        CO_E_FAILEDTOOPENPROCESSTOKEN = 0x8001013C,
        CO_E_DECODEFAILED = 0x8001013D,
        CO_E_ACNOTINITIALIZED = 0x8001013F,
        CO_E_CANCEL_DISABLED = 0x80010140,
        RPC_E_UNEXPECTED = 0x8001FFFF,
    }
}