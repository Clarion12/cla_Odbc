
!ABCIncludeFile

  OMIT('_EndOfInclude_',_ODBCConnectionPresent_)
_ODBCConnectionPresent_   EQUATE(1)

  include('odbcTypes.inc'),once
  include('dynstr.inc'),once
  include('odbcConnStrCl.inc'),once
  include('OdbcHandleCl.inc'),once
                       
  ITEMIZE(0),PRE(Connection)
Opened equate
CallerOpened  equate
failed equate
  end
  
ODBCConnectionClType class,type,module('odbcConn.clw'),link('odbcConn.clw',_ABCLinkMode_),dll(_ABCDllMode_)
! ODBC handles 
hDbc                   &OdbcHandleClType,private ! db connection handle
hEnv                   &OdbcHandleClType,private ! enviorment handle
hStmt                  &OdbcStmtHandleClType,private! statement handle
!
connStr                &baseConnStrClType,private ! instance to hold the connection string
! -------------------------------------------------------
! some getters, no setters.  code will need to query these values
! and use locally, but they will never be modifed
! ------------------------------------------------------
gethDbc                procedure(),SQLHDBC
gethStmt               procedure(),SQLHStmt
gethEnv                procedure(),SQLHEnv

! ------------------------------------------------------
! checks to see if the connection handle input or the 
! instance connection handle is active or dead
! ------------------------------------------------------
isConnected            procedure(SQLHDBC dbHandle = 0),bool

! set up and clean up
init                   procedure(),sqlReturn
init                   procedure(baseConnStrClType  connString),sqlReturn,proc
kill                   procedure()

! ------------------------------------------------------
! sets the ODBC version in use 
! ------------------------------------------------------
setOdbcVersion         procedure(long verId),sqlReturn,proc 

! ------------------------------------------------------
! connect to the database
! if withStatement is true, the default.  a statment handle is 
! allocated 
! ------------------------------------------------------
Connect                procedure(bool statement = withStatement),sqlReturn

! ------------------------------------------------------
! disconnect from the database
! if with statment handle is true the handel is released
! ------------------------------------------------------
Disconnect             procedure(bool statement = withStatement),sqlReturn,proc

! read the errors and show if wanted
getError               procedure(SQLSMALLINT HandleType, SQLHANDLE Handle),virtual,protected
getDatabaseError       procedure(),virtual,protected

asyncOn procedure(),sqlReturn

                     end ! class 


  _EndOfInclude_
