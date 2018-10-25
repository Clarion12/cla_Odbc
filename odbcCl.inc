

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
conn             &ODBCConnectionClType,private
! set up and tear down
init             procedure(*ODBCConnectionClType conn),byte,proc
kill             procedure(),virtual
destruct         procedure()
! public wirtuals 
formatRow        procedure(),virtual
! public protected
setSqlCommand    procedure(*IdynStr s),sqlReturn,protected
setSqlCommand    procedure(string s),sqlReturn,protected
!
nextResultSet    procedure(),bool
fetch            procedure(*queue q),sqlReturn,private
fetchPage        procedure(long sqlDir),sqlReturn,proc
closePaging      procedure(),sqlReturn,proc
fillResult       procedure(*columnsClass cols, *queue q),sqlReturn,private
getError         procedure(),private
! public methods that are used to run a query 
prepQuery       procedure(*IDynStr sqlCode),sqlReturn,virtual
execQuery       procedure(*IDynStr sqlCode),sqlReturn,virtual
execQuery       procedure(*IDynStr sqlCode, *columnsClass cols, *queue q),sqlReturn,virtual
execQuery       procedure(*sqlStrClType sqlCode, *columnsClass cols, *queue q),sqlReturn,virtual
execQuery       procedure(*IDynStr sqlCode, *columnsClass cols, *ParametersClass params, *queue q),sqlReturn,virtual
! private query 
execQuery       procedure(),sqlReturn,private
execQuery       procedure(*ParametersClass params),sqlReturn,private
setupQuery      procedure(*IDynStr sqlCode, *columnsClass cols),sqlReturn,private
!
! code to call stored procedures 
execSp         procedure(string spName, *ParametersClass params),sqlReturn
execSp         procedure(string spName, *columnsClass cols, *queue q),sqlReturn
execSp         procedure(string spName, *columnsClass cols, *ParametersClass params, *queue q),sqlReturn
! call a sclar function
callScalar     procedure(string spName, *ParametersClass params),sqlReturn 
! private
execSp         procedure(),private,sqlReturn
setupSpCall   procedure(string spName),sqlReturn,private
setupSpCall   procedure(string spName, *ParametersClass params),sqlReturn,private
            end

  _EndOfInclude_
