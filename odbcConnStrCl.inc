
!ABCIncludeFile

  OMIT('_EndOfInclude_',_ODBCConnectionStrPresent_)
_ODBCConnectionStrPresent_   EQUATE(1)

  include('odbcTypes.inc'),once
  include('dynstr.inc'),once

! some default string lengths of the connections   
eDriverNameLength equate(50)  
! Oracle limits to 30 characaters, typically not an issue with other vendors.
eDefaultLength    equate(30)

! ---------------------------------------------------------------------------
! interface for the connection strings.  While many of the properties are the same
! each vendor has some differences.  Use this base interface for the typical items and 
! derive the interface for vendor specifics, ie MARS with MS-SQL Server
! ---------------------------------------------------------------------------
IConnString       interface 
kill                procedure()
! some setters for the fields, no getters
setDriverName       procedure(string driverName)
setDbName           procedure(string dbname)
setSrvName          procedure(string srvName)
setUserName         procedure(string  user)
setPassword         procedure(string pw)
setPortNumber       procedure(string portNumber)
setTrustedConn      procedure(bool onOff)
! always implement this method, the actual format of the connection 
! string varies significantly
ConnectionString    procedure(),*cstring
                  end 
! ---------------------------------------------------------------------------

! ---------------------------------------------------------------------------
! baseConnStrClType, class to handle the standard connection strings.  Each 
! vendor will use this as the starting point, derive the class for any specifics.
! 
! ---------------------------------------------------------------------------
baseConnStrClType  class,type,module('odbcConnStrCl.clw'),link('odbcConnStrCl.clw',_ABCLinkMode_),dll(_ABCDllMode_)
connStr              &IDynStr,private
IConnStr             &IConnString
! connection string fields, access is through the interface 
srvName              cstring(eDefaultLength),private
dbName               cstring(eDefaultLength),private
driverName           cstring(eDriverNameLength),private
userName             cstring(eDefaultLength),private
password             cstring(eDefaultLength),private
portNumber           cstring(eDefaultLength),private
trustedConn          bool
! setup and clean up, 
! set the type and then use the interface for the fields.
init                 procedure(long db),*IConnString,virtual
! pretty much all the vendors use these three strings, use the setters for any additional properties
init                 procedure(long db, string driverName, string srvName, string dbName),*IConnString,virtual
kill                 procedure(),virtual
                   end
! ---------------------------------------------------------------------------

! ---------------------------------------------------------------------------
! derive the interface and class for MS, add the MARS option to the connection
! string.
! ---------------------------------------------------------------------------
ISqlConnStr interface(IConnString)
setUseMars    procedure(bool onOff)
           end 
                
MSConnStrClType  class(baseConnStrClType),type,implements(ISqlConnStr),module('odbcConnStrCl.clw'),link('odbcConnStrCl.clw',_ABCLinkMode_),dll(_ABCDllMode_)
useMars            bool,private
                 end 
! ---------------------------------------------------------------------------

! ---------------------------------------------------------------------------
! derive the interface and class for PostGreSql, no additional fields or propeties 
! but go ahead and derive, on the small chance something needs to be added and 
! mainly to give each vendors connection string instances the same interface
! ---------------------------------------------------------------------------
IPgConnStr   interface(IConnString)
             end 
            
pgConnStrClType  class(baseConnStrClType),type,implements(IPgConnStr),module('odbcConnStrCl.clw'),link('odbcConnStrCl.clw',_ABCLinkMode_),dll(_ABCDllMode_)
                 end 
! ---------------------------------------------------------------------------
                   
! ---------------------------------------------------------------------------
! derive the interface and class for oracle, no additional fields or propeties 
! but go ahead and derive, on the chance something needs to be added and 
! mainly to give each vendors connection string instances the same interface
! ---------------------------------------------------------------------------
IOraConnStr   interface(IConnString)
             end 
            
oraConnStrClType class(baseConnStrClType),type,implements(IOraConnStr),module('odbcConnStrCl.clw'),link('odbcConnStrCl.clw',_ABCLinkMode_),dll(_ABCDllMode_)
                 end 
! ---------------------------------------------------------------------------                   

   _EndOfInclude_                   
