
!ABCIncludeFile

  OMIT('_EndOfInclude_',_ODBCColumssPresent_)
_ODBCColumssPresent_   EQUATE(1)

  include('odbcTypes.inc'),once
  include('odbcHandleCl.inc'),once
  include('dynstr.inc'),once
  
! queue to store the columns used in a query
ColumnsQueue queue,type
ColId          long           ! ord of the column in the select statement
ColName        string(eColNameLength)  ! name of the col, not used by the odbc but by the classes
ColType        SQLSMALLINT    ! type of the col, odbc const
ColValue       SQLPOINTER     ! pointer to the buffer for the result
ColSize        SQLLEN         ! size of the column in bytes
resultSetId    long
             end 

columnsClass class,type,module('odbcColumnsCl.clw'),link('odbcColumnsCl.clw',_ABCLinkMode_),dll(_ABCDllMode_)
colQ           &ColumnsQueue,private
fieldList      &IDynStr
!
construct      procedure()
init           procedure(),byte,proc
kill           procedure()
destruct       procedure()
clearQ         procedure()

bindColumns    procedure(long  hStmt),sqlReturn,virtual
getFields      procedure(),*cstring,virtual
!
AddColumn      procedure(*long colPtr),sqlReturn,proc
AddColumn      procedure(string colLabel, *long colPtr),sqlReturn,proc
!
AddColumn      procedure(*string colPtr),sqlReturn,proc     
AddColumn      procedure(string colLabel, *string colPtr),sqlReturn,proc     
!
AddColumn      procedure(*real colPtr),sqlReturn,proc
AddColumn      procedure(string colLabel, *real colPtr),sqlReturn,proc
!
AddColumn      procedure(*TIMESTAMP_STRUCT colPtr),sqlReturn,proc
AddColumn      procedure(string colLabel, *TIMESTAMP_STRUCT colPtr),sqlReturn,proc
!
AddColumn      procedure(SQLSMALLINT TargetType, SQLPOINTER TargetValuePtr, SQLLEN BufferLength),sqlReturn,private,proc
!
addField       procedure(string colLabel),protected,virtual
            end 

  _EndOfInclude_
