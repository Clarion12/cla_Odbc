

!ABCIncludeFile

  OMIT('_EndOfInclude_',_ODBCClassPresent_)
_ODBCClassPresent_   EQUATE(1)

  include('dynStr.inc'),once
  include('odbcConn.inc'),once
  include('odbcBaseCl.inc'),once
  include('odbcTypes.inc'),once
  include('odbcParamsCl.inc'),once 
  include('odbcSqlStrCl.inc'),once
  include('odbcColumnsCl.inc'),once 

! ----------------------------------------------------------------------
! class that is used to execute a query from the file manage rinstance. 
! the query is passed in a string and then formated and called.
! 
! result sets, if any, are processed.  input and output parameters 
! can be used. 
!
! Note, the term exec in the name does not imply that the query is caled using the 
! exec 'some statement' format.  All these calls use the call syntax.
! the call syntax does less work than the exec style.
! ----------------------------------------------------------------------
odbcExecType class(odbcBaseClType),type,module('odbcCl.clw'),link('odbcCl.clw',_ABCLinkMode_),dll(_ABCDllMode_)

! set up and tear down
! ----------------------------------------------------------------------
! initilizes the object 
! ----------------------------------------------------------------------
init             procedure(),byte,proc


destruct         procedure(),virtual

! ----------------------------------------------------------------------
! virtual place holder
! use this function to format the fields or columns read prior to the display
! ----------------------------------------------------------------------
formatRow        procedure(),virtual

! ----------------------------------------------------------------------
! execute a query that does not return a result set 
! ----------------------------------------------------------------------
execQuery       procedure(SQLHSTMT hStmt, *IDynStr sqlCode),sqlReturn,virtual

! ------------------------------------------------------------------------------
! execQuery
! execute a query that returns a result set.  
! prep the statement, execute the statement
! then fill the queue or buffers and close the connection when done
!
! this method is used for qureries that do not accept parameters 
! ------------------------------------------------------------------------------    
execQuery       procedure(SQLHSTMT hStmt, *IDynStr sqlCode, *columnsClass cols, *queue q),sqlReturn,virtual

! ------------------------------------------------------------------------------
! execQuery
! execute a query that returns a result set.  
! prep the statement, execute the statement
! then fill the queue or buffers and close the connection when done
!
! this method is used for qureries that do not accept parameters 
! ------------------------------------------------------------------------------    
execQuery       procedure(SQLHSTMT hStmt, *IDynStr sqlCode, *columnsClass cols, *ParametersClass params, *queue q),sqlReturn,virtual

! ------------------------------------------------------------------------------
! execQuery
! execute a query that does not return a result set.  
! but has at least one output parameter
! ------------------------------------------------------------------------------    
execQueryOut    procedure(SQLHSTMT hStmt, *IDynStr sqlCode, *ParametersClass params),sqlReturn,virtual

! ------------------------------------------------------------------------------
! execQuery
! execute a query that does not return a result set.  
! but has input parameters and may have output parameters.
! ------------------------------------------------------------------------------      
execQuery       procedure(SQLHSTMT hStmt, *IDynStr sqlCode, *ParametersClass params),sqlReturn,virtual

            end

  _EndOfInclude_
