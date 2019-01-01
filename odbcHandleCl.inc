
!ABCIncludeFile

  OMIT('_EndOfInclude_',_ODBCHandlePresent_)
_ODBCHandlePresent_   EQUATE(1)

  include('odbcTypes.inc'),once
  include('odbcErrorCl.inc'),once  
                       
! ------------------------------------------------------------------
! base class for the three handle types used in the calls
! hEnv, hDbc and hStmt 
! allocates a handle and frees a handle there is a getter function 
! to get the value of the handle, there is no setter 
! ------------------------------------------------------------------                       
OdbcHandleClType class,type,module('odbcHandleCl.clw'),link('odbcHandleCl.clw',_ABCLinkMode_),dll(_ABCDllMode_)
handle             SQLHANDLE,private
handleType         SQL_HANDLE_TYPE,private
kill               procedure(),virtual
destruct           procedure()
allocateHandle     procedure(SQL_HANDLE_TYPE hType, SQL_HANDLE_TYPE pType),sqlReturn,proc,virtual
freeHandle         procedure(),sqlReturn,proc,virtual
getHandle          procedure(),SQLHANDLE
                 end

! ------------------------------------------------------------------
! derived class for the statement handle 
! allocates a handle and frees a handle there is a getter function 
! to get the value of the handle, there is no setter 
!
! a derived object is used for the statment handle.  there are 
! cases where a statement handle can be reused so this object 
! provides a function to clear any bindings from the handle 
! and allows the application to allocate more than one statment handle 
! when using MARS
! ------------------------------------------------------------------                       
OdbcStmtHandleClType class(OdbcHandleClType),type,module('odbcHandleCl.clw'),link('odbcHandleCl.clw',_ABCLinkMode_),dll(_ABCDllMode_)

! ------------------------------------------------------------
! unbinds the columns and the parameters for the state,ent 
! handle.  the statment handle at this stime could be reused 
! if needed/wanted
! ------------------------------------------------------------
freeBindings           procedure(),sqlReturn,proc,virtual

! ------------------------------------------------------------
! unbindsreleases the statement handle and removes all bindings
! ------------------------------------------------------------
freeHandle             procedure(),sqlReturn,proc,virtual
                     end                 
                
  _EndOfInclude_