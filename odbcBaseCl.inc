!ABCIncludeFile

  OMIT('_EndOfInclude_',_ODBCBaseClassPresent_)
_ODBCBaseClassPresent_   EQUATE(1)

  include('dynStr.inc'),once
  include('odbcTypes.inc'),once
  include('odbcParamsCl.inc'),once 
  include('odbcSqlStrCl.inc'),once
  include('odbcColumnsCl.inc'),once 


odbcBaseClType class,type,module('odbcBaseCl.clw'),link('odbcbaseCl.clw',_ABCLinkMode_),dll(_ABCDllMode_)

! -----------------------------------------------------------
! string class is used to build the calls to tstored procedures 
! and scalar functions.  adds the {call ...} and adds the palces 
! holders for each parameter 
! not used when executing a query
! -----------------------------------------------------------
sqlStr           &sqlStrClType

! set up and tear down
init             procedure(*ODBCErrorClType e),byte,proc
kill             procedure(),virtual
destruct         procedure(),virtual

! ------------------------------------------------------
! instance of the error class for the ODBC messages
! this is private onject.  there is a function that will
! allow error class message queue to be queried.
! ------------------------------------------------------
errs                   &ODBCErrorClType,private 

! ----------------------------------------------------------------------
! unbinds any columns that are currently bound
! typical use is when a call that returns multiple result sets is in use.
! the first result set is processed and then call this to unbind the columns,
! then bind the caloums for the second result set, 
! repeat as needed
! note, this call does no free the statement handle
! ----------------------------------------------------------------------
unBindColums     procedure(SQLHSTMT hStmt)  

! ----------------------------------------------------------------------
! setss the isolation level for the hDbc input.  the connection must not 
! have any open transaction when this is called.  the connection can be opend 
! or closed. 
! ----------------------------------------------------------------------
setIsolationLevel procedure(SQLHDBC hDbc, long level),sqlReturn

! ----------------------------------------------------------------------
! begins a transaction for the connection handle input.  
! this actually turns off auto-commit mode.
! ----------------------------------------------------------------------
beginTransaction procedure(SQLHDBC hDbc),sqlReturn

! ----------------------------------------------------------------------
! commits a transaction for the connection handle input.  
! ----------------------------------------------------------------------
CommitTransaction procedure(SQLHDBC hDbc),sqlReturn

! ----------------------------------------------------------------------
! rolls back a transaction for the connection handle input.  
! ----------------------------------------------------------------------
RollbackTransaction procedure(SQLHDBC hDbc),sqlReturn

! ----------------------------------------------------------------------
! ends a transaction for the connection handle input.  
! called fromthe commit or rollback functions.
! ----------------------------------------------------------------------
EndTransaction procedure(SQLHDBC hDbc, long committRollBack),sqlReturn,private

formatRow        procedure(),virtual  

! -----------------------------------------------------------------------------
! worker function to assign the sql str (the actual sql statement) used in this 
! call to the class member, the string parameter is typically used by the various
! calls to stored procedures
! -----------------------------------------------------------------------------  
setSqlCommand    procedure(string s),sqlReturn,protected

! ----------------------------------------------------------------------  
! checks for the next result set, if any, and moves to the next result set
! returns true if there is more and false if not
! ----------------------------------------------------------------------  
nextResultSet    procedure(SQLHSTMT hStmt),bool

! ------------------------------------------------------------------------------
! reads the next result set by calling the fetch function
! ------------------------------------------------------------------------------
readnextResult   procedure(SQLHSTMT hStmt, *queue q),sqlReturn

executeDirect    procedure(SQLHSTMT hStmt, *IDynStr sqlCode),sqlReturn,virtual,protected

! ------------------------------------------------------------------------------
! fetch with out a result set.  
! ------------------------------------------------------------------------------
fetch            procedure(SQLHSTMT hStmt),sqlReturn,virtual

! ------------------------------------------------------------------------------
! fetch
! reads the result set, one row at a time and places the data into the queue fields.
! Queue fields are already bound to the columns so all that is needed here is an add(q)
!
! Note, the queue fields must be bound before this method is called.
! ------------------------------------------------------------------------------
fetch            procedure(SQLHSTMT hStmt, *queue q),sqlReturn,protected,virtual

fetch            procedure(SQLHSTMT hStmt, *queue q, columnsClass cols),sqlReturn,protected,virtual

! -----------------------------------------------------------------------------
! Binds the columns from the queue to the columns in the result set
! then calls fetch to read the result set
! -----------------------------------------------------------------------------
fillResult       procedure(SQLHSTMT hStmt, *columnsClass cols, *queue q, long setId = 1),sqlReturn,protected

! -----------------------------------------------------------------------------
! process any large character or binay objects in the query
! -----------------------------------------------------------------------------
processLargeObjects procedure(SQLHSTMT hStmt, *columnsClass cols)

! -----------------------------------------------------------------------------
! reads the large character objects 
! -----------------------------------------------------------------------------
readLargeChar       procedure(SQLHSTMT hStmt, *columnsClass cols) 

! -----------------------------------------------------------------------------
! reads the number of bytes for the large character objects 
! -----------------------------------------------------------------------------
getCharLength       procedure(SQLHSTMT hStmt, *columnsClass cols),long,private 

! -----------------------------------------------------------------------------
! reads the large binary objects 
! -----------------------------------------------------------------------------
readLargeBinary     procedure(SQLHSTMT hStmt, *columnsClass cols) 

! -----------------------------------------------------------------------------
! reads the number of bytes for the large binary objects 
! -----------------------------------------------------------------------------
getBinaryLength     procedure(SQLHSTMT hStmt, *columnsClass cols),long,private 


ProcessRead      procedure(*queue q, *columnsClass cols),virtual,protected

! -----------------------------------------------------------------------------
! call the error class to read the error information
! -----------------------------------------------------------------------------
getError         procedure(SQLHSTMT hStmt),protected

! -----------------------------------------------------------------------------
! call the error class to read the error information
! -----------------------------------------------------------------------------
getConnectionError procedure(SQLHDBC hDbc),protected

            end

  _EndOfInclude_
