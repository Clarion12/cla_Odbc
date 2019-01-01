

!ABCIncludeFile

  OMIT('_EndOfInclude_',_ODBCClassPresent_)
_ODBCClassPresent_   EQUATE(1)

  include('dynStr.inc'),once
  include('odbcConn.inc'),once
  include('odbcTypes.inc'),once
  include('odbcParamsCl.inc'),once 
  include('odbcSqlStrCl.inc'),once
  include('odbcColumnsCl.inc'),once 


odbcClType class,type,module('odbcCl.clw'),link('odbcCl.clw',_ABCLinkMode_),dll(_ABCDllMode_)
sqlStr           &sqlStrClType
conn             &ODBCConnectionClType,protected
! set up and tear down
init             procedure(*ODBCConnectionClType conn),byte,proc
kill             procedure(),virtual
destruct         procedure(),virtual

unBindColums     procedure()  

! ----------------------------------------------------------------------
! virtual place holder
! use this function to format the fields or columns read prior to the display
! ----------------------------------------------------------------------
formatRow        procedure(),virtual

! -----------------------------------------------------------------------------
! worker function to assign the sql str (the actual sql statement) used in this 
! call to the class member, the IDynStr is typically used by the various execute
! calls
! -----------------------------------------------------------------------------  
setSqlCommand    procedure(*IdynStr s),sqlReturn,protected

! -----------------------------------------------------------------------------
! worker function to assign the sql str (the actual sql statement) used in this 
! call to the class member, the string parameter is typically used by the various
! execute calls to stored procedures
! -----------------------------------------------------------------------------  
setSqlCommand    procedure(string s),sqlReturn,protected

! ----------------------------------------------------------------------  
! checks for the next result set, if any, and moves to the next result set
! returns true if there is more and false if not
! ----------------------------------------------------------------------  
nextResultSet    procedure(),bool

! ------------------------------------------------------------------------------
! reads the next result set by calling the fetch function
! ------------------------------------------------------------------------------
readnextResult   procedure(*queue q),sqlReturn

! ------------------------------------------------------------------------------
! fetch with out a result set.  
! ------------------------------------------------------------------------------
fetch            procedure(),sqlReturn,virtual

! ------------------------------------------------------------------------------
! fetch
! reads the result set, one row at a time and places the data into the queue fields.
! Queue fields are already bound to the columns so all that is needed here is an add(q)
!
! Note, the queue fields must be bound before this method is called.
! ------------------------------------------------------------------------------
fetch            procedure(*queue q),sqlReturn,protected,virtual

! -----------------------------------------------------------------------------
! Binds the columns from the queue to the columns in the result set
! then calls fetch to read the result set
! -----------------------------------------------------------------------------
fillResult       procedure(*columnsClass cols, *queue q, long setId = 1),sqlReturn,protected

! -----------------------------------------------------------------------------
! call the error class to read the error information
! -----------------------------------------------------------------------------
getError         procedure(),private

! ----------------------------------------------------------------------
! execute a query that does not return a result set 
! ----------------------------------------------------------------------
execQuery       procedure(*IDynStr sqlCode),sqlReturn,virtual

! ------------------------------------------------------------------------------
! execQuery
! execute a query that returns a result set.  
! prep the statement, execute the statement
! then fill the queue or buffers and close the connection when done
!
! this method is used for qureries that do not accept parameters 
! ------------------------------------------------------------------------------    
execQuery       procedure(*IDynStr sqlCode, *columnsClass cols, *queue q),sqlReturn,virtual

! ------------------------------------------------------------------------------
! execQuery
! execute a query that returns a result set.  
! prep the statement, execute the statement
! then fill the queue or buffers and close the connection when done
!
! this method is used for qureries that do not accept parameters 
! ------------------------------------------------------------------------------    
execQuery       procedure(*IDynStr sqlCode, *columnsClass cols, *ParametersClass params, *queue q),sqlReturn,virtual

! ------------------------------------------------------------------------------
! execQuery
! execute a query that does not return a result set.  
! prep the statement, execute the statement
! then fill the queue or buffers and close the connection when done
! ------------------------------------------------------------------------------    
execQueryOut    procedure(*IDynStr sqlCode, *ParametersClass params),sqlReturn,virtual

! ------------------------------------------------------------------------------    
! execute the query, this call the ODBC SQLExecDirect and is called internally 
! by the other execute functions
! ------------------------------------------------------------------------------    
execQuery       procedure(),sqlReturn,private

! ------------------------------------------------------------------------------    
! binds the parameters just before the execution
! ------------------------------------------------------------------------------    
execQuery       procedure(*ParametersClass params),sqlReturn,private

setupQuery      procedure(*IDynStr sqlCode, *columnsClass cols),sqlReturn,private
!

! ------------------------------------------------------------------
! call stored procedure with one or more table valued parameters.
! this is a virtual place holder and needs to be overloaded in function
! or in some code that can hold the array's until the write completes.
! ------------------------------------------------------------------
execTableSp    procedure(string spName, *ParametersClass params, long numberRows),sqlReturn,virtual

execSp         procedure(string spName),sqlReturn,virtual
execSp         procedure(string spName, *ParametersClass params),sqlReturn,virtual
execSp         procedure(string spName, *columnsClass cols, *queue q),sqlReturn,virtual
execSp         procedure(string spName, *columnsClass cols, *ParametersClass params, *queue q),sqlReturn,virtual
! call a sclar function
callScalar     procedure(string spName, *ParametersClass params),sqlReturn 
! private
execSp         procedure(),sqlReturn,virtual
setupSpCall   procedure(string spName),sqlReturn,virtual
setupSpCall   procedure(string spName, *ParametersClass params),sqlReturn,virtual
            end

  _EndOfInclude_
