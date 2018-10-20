
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
bindParameters        procedure(SQLHSTMT hStmt),sqlReturn
bindParameters        procedure(SQLHSTMT hStmt, *sqlStrClType sqlStr),sqlReturn
AddInParameter        procedure(*long colPtr),sqlReturn,proc
AddInParameter        procedure(string pName, *long varPtr),sqlReturn,proc
AddInParameter        procedure(*string colPtr),sqlReturn,proc     
AddInParameter        procedure(string pName, *string colPtr),sqlReturn,proc     
AddInParameter        procedure(*real colPtr),sqlReturn,proc
AddInParameter        procedure(string pname, *real colPtr),sqlReturn,proc
AddInParameter        procedure(*TIMESTAMP_STRUCT colPtr),sqlReturn,proc
AddInParameter        procedure(string pname, *TIMESTAMP_STRUCT colPtr),sqlReturn,proc
addParameter          procedure(SQLSMALLINT InOutType, SQLSMALLINT ValueType, SQLSMALLINT ParameterType, SQLULEN ColumnSize, SQLSMALLINT DecimalDigits, SQLPOINTER varPtr, SQLLEN BufferLength),byte,proc,private
addParameter          procedure(string paramName, SQLSMALLINT InOutType, SQLSMALLINT ValueType, SQLSMALLINT ParameterType, SQLULEN ColumnSize, SQLSMALLINT DecimalDigits, SQLPOINTER varPtr, SQLLEN BufferLength),byte,proc,private
!
FillPlaceHolders      procedure(sqlStrClType sqlCode),long
findByParamName       procedure(string pName),sqlReturn
                    end

  _EndOfInclude_
