
!ABCIncludeFile

  OMIT('_EndOfInclude_',_ODBCColumssPresent_)
_ODBCColumssPresent_   EQUATE(1)

  include('odbcTypes.inc'),once
  include('odbcHandleCl.inc'),once
  include('dynstr.inc'),once

! -----------------------------------------------------------  
! queue to store the columns used in a query
! -----------------------------------------------------------  
ColumnsQueue queue,type
ColId          long           ! ord of the column in the select statement
ColName        string(eColNameLength)  ! name of the col, not used by the odbc but by the classes
ColType        SQLSMALLINT    ! type of the col, odbc const
ColValue       SQLPOINTER     ! pointer to the buffer for the result
ColSize        SQLLEN         ! size of the column in bytes
             end 

! --------------------------------------------------------------------- 
! class to manage the colums used in a query or a stored procedure call
! each column used is added to the queue and will be mapped or bound 
! to the destination.  the destenation can be any buffer needed/wanted,
! a queue, a group, an individual field, ...
! --------------------------------------------------------------------  
columnsClass class,type,module('odbcColumnsCl.clw'),link('odbcColumnsCl.clw',_ABCLinkMode_),dll(_ABCDllMode_)
! queue of the columns that will be used by the databse call
colQ           &ColumnsQueue,private
! list of the fields in a comma delimited string
! not yet implemented, will be used to build add hoc sql statements
fieldList      &IDynStr

! ---------------------------------------------------------------------------
!  default constructor, calls the init function for the set up
! ---------------------------------------------------------------------------
construct      procedure()

! ---------------------------------------------------------------------------
!  allocates the queue and the dyn str used
! ---------------------------------------------------------------------------  
init           procedure(),byte,proc

! ------------------------------------------------------------------------------
! disposes the queue and the dyn str
! ------------------------------------------------------------------------------
kill           procedure()

! ---------------------------------------------------------------------------
!  default destructor, calls the kill function for the clean up
! ---------------------------------------------------------------------------
destruct       procedure(),virtual

! ------------------------------------------------------------------------------
! free the queue and the dyn str that contains the fileds names
! ------------------------------------------------------------------------------
clearQ         procedure()

! ------------------------------------------------------------------------------
! gets the string or the list of column labels stored in the dyn str, 
! returns a cstring, not implemented
! ------------------------------------------------------------------------------
getFields      procedure(),*cstring,virtual

! ------------------------------------------------------------------------------
! Bind the queue, group or seperate fields to the columns in the result set.
! column order is typically the  same order as the select statment,
! 
! parameters for the ODBC api call 
! hStmt   = handle to the ODBC statement
! colId = ord value of the parmaeter, 1, 2, 3 ... the ordinal position
! colType = the C data type of the column 
! colValue = pointer to the buffer field 
! colSize = the size of the buffer or the queue field 
! colInd = pointer to a buffer for the size of the parameter. not used and null in this example 
! -----------------------------------------------------------------------------    
bindColumns    procedure(long  hStmt),sqlReturn,virtual

! ------------------------------------------------------------------------------
! the various addColumn functions are called by the using code and are used for the
! specific data types.  each calls the AddColumn/3 function to actually 
! add a columns
! ------------------------------------------------------------------------------
AddColumn      procedure(*long colPtr)
AddColumn      procedure(string colLabel, *long colPtr)
!
AddColumn      procedure(*string colPtr)
AddColumn      procedure(string colLabel, *string colPtr)
!
AddColumn      procedure(*real colPtr)
AddColumn      procedure(string colLabel, *real colPtr)
!
AddColumn      procedure(*TIMESTAMP_STRUCT colPtr)
AddColumn      procedure(string colLabel, *TIMESTAMP_STRUCT colPtr)

! ------------------------------------------------------------------------------
! add a column to the queue.  each column added here will be bound to the 
! statment handle.  The colums are bound when the execute function is called.
! this is typically called by the various addColumn/1 functions but can be called directly
! in a derived instance.
!
! the columns are boud in the sequence they are added.
! ------------------------------------------------------------------------------  
AddColumn      procedure(SQLSMALLINT TargetType, SQLPOINTER TargetValuePtr, SQLLEN BufferLength),protected

! ---------------------------------------------------------
! adds a label to the list of columns not implemented
! ---------------------------------------------------------
addField       procedure(string colLabel),protected,virtual
            end 

  _EndOfInclude_
