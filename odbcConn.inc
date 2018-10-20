
!ABCIncludeFile

  OMIT('_EndOfInclude_',_ODBCConnectionPresent_)
_ODBCConnectionPresent_   EQUATE(1)

  include('odbcTypes.inc'),once
  include('dynstr.inc'),once
  !include('odbcErrorCl.inc'),once  
  include('odbcConnStrCl.inc'),once
  include('OdbcHandleCl.inc'),once
                       

ODBCConnectionClType class,type,module('odbcConn.clw'),link('odbcConn.clw',_ABCLinkMode_),dll(_ABCDllMode_)
! ODBC handles 
hDbc                   &OdbcHandleClType,private ! db connection handle
hEnv                   &OdbcHandleClType,private ! db enviorment handle
hStmt                  &OdbcStmtHandleClType,private! db statement handle
!
connStr                &IConnString,private ! instance to hold the connection string
! some getters, no setters.  Sometimes code will need to query these values
! and use locally, but no code should be allowed to modify
gethDbc                procedure(),SQLHDBC
gethStmt               procedure(),SQLHStmt
gethEnv                procedure(),SQLHEnv
!
init                   procedure(),sqlReturn
init                   procedure(*IConnString connString),sqlReturn,proc
kill                   procedure()
! public methods 
setOdbcVersion         procedure(),sqlReturn,proc 
connect                procedure(bool statement = withStatement),sqlReturn,proc
Disconnect             procedure(bool statement = withStatement),sqlReturn,proc
! 
getError               procedure(SQLSMALLINT HandleType, SQLHANDLE Handle),virtual,protected
getDatabaseError       procedure(),virtual,protected

                     end ! class 


  _EndOfInclude_
