
!ABCIncludeFIle

  OMIT('_EndOfInclude_',_ODBCParamPresent_)
_ODBCParamPresent_   EQUATE(1)

  include('odbcTypes.inc'),once
  include('dynstr.inc'),once
  include('odbcSqlStrCl.inc'),once
    
! queue to store information for ech parameter used in a sp call
! or in a query that accepts parameters
ParametersQueue queue,type
paramId           long    
paramName         string(eParamNameLength)      
InOutType         SQLSMALLINT   
valueType         SQLSMALLINT   
paramType         SQLSMALLINT   
paraSize          SQLULEN       
DecimalDigits     SQLSMALLINT   
ParamPtr          SQLPOINTER    
paramLength       SQLLEN
tableName         cstring(128)
                end ! queue 
                
!
ParametersClass class,type,module('odbcParamsCl.clw'),link('odbcParamsCl.clw',_ABCLinkMode_),dll(_ABCDllMode_)
!bindByName            bool,private
setupFailed           bool,private
paramQ                &ParametersQueue
construct             procedure()
init                  procedure(),byte,proc
kill                  procedure()
destruct              procedure()
clearQ                procedure()
!
bindParameters        procedure(SQLHSTMT hStmt, *long numberRows),sqlReturn
bindParameters        procedure(SQLHSTMT hStmt),sqlReturn
focusTableParameter   procedure(SQLHSTMT hStmt, long ordinal)  
unfocusTableParameter procedure(SQLHSTMT hStmt)
! ----------------------------------------------------------
! add the various input parameters
! ----------------------------------------------------------

AddInParameter        procedure(*long colPtr),sqlReturn,proc
AddInParameter        procedure(string pName, *long varPtr),sqlReturn,proc
AddInParameter        procedure(*cstring colPtr),sqlReturn,proc     
AddInParameter        procedure(string pName, *cstring colPtr),sqlReturn,proc     
AddInParameter        procedure(*real colPtr),sqlReturn,proc
AddInParameter        procedure(string pname, *real colPtr),sqlReturn,proc
AddInParameter        procedure(*TIMESTAMP_STRUCT colPtr),sqlReturn,proc
AddInParameter        procedure(string pname, *TIMESTAMP_STRUCT colPtr),sqlReturn,proc
AddTableParameter     procedure(long numberRows, *cstring tableName),sqlReturn,proc
addParameter          procedure(SQLSMALLINT InOutType, SQLSMALLINT ValueType, SQLSMALLINT ParameterType, SQLULEN ColumnSize, SQLSMALLINT DecimalDigits, SQLPOINTER varPtr, SQLLEN BufferLength),byte,proc
addParameter          procedure(string paramName, SQLSMALLINT InOutType, SQLSMALLINT ValueType, SQLSMALLINT ParameterType, SQLULEN ColumnSize, SQLSMALLINT DecimalDigits, SQLPOINTER varPtr, SQLLEN BufferLength),byte,proc
! ----------------------------------------------------------------
! add the output types
! ----------------------------------------------------------------
AddOutParameter       procedure(*long varPtr),sqlReturn,proc
AddOutParameter       procedure(*cstring varPtr),sqlReturn,proc
! 
FillPlaceHolders      procedure(sqlStrClType sqlCode, long startPos = 1),long
findByParamName       procedure(string pName),sqlReturn
                    end

  _EndOfInclude_
