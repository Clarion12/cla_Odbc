
  member()
  
  include('odbcHandleCl.inc'),once 
  include('odbcTypes.inc'),once 

  map 
    module('odbc32')
      SQLAllocHandle(SQLSMALLINT HandleType, SQLHANDLE InputHandle, *SQLHANDLE OutputHandlePtr),SqlReturn,pascal
      SQLFreeHandle(SqlSmallInt hType, SqlHandle h),long,pascal
      SQLFreeStmt(SQLHSTMT StatementHandle, SQLSMALLINT opt),sqlReturn,pascal,proc
      SQLSetEnvAttr(SQLHENV EnvironmentHandle, SQLINTEGER Attribute,  SQLPOINTER Value, SQLINTEGER StringLength),sqlReturn,pascal
    end 
  end

OdbcHandleClType.kill procedure() !,virtual

retv   sqlReturn,auto

  code 
  
  retv = self.freeHandle()    
  if (retv = sql_Success) 
    self.handle = 0  
  end 
  
  return 
! end kill 
! ----------------------------------------------------------------------------
    
OdbcHandleClType.destruct procedure()

  code 

  self.kill() 
  
  return

OdbcHandleClType.allocateHandle procedure(SQL_HANDLE_TYPE hType, SQL_HANDLE_TYPE pType) !,sqlReturn,proc

err ODBCErrorClType

retv   sqlReturn,auto

verId  ulong(SQL_OV_ODBC3_80)

  code
  
  self.handleType = hType
  retv = SQLAllocHandle(hType, pType, self.handle)
  
  if (retv <> Sql_Success) 
    err.getError(pType, self.handle)
  else 
    ! set the version attribute if this is an SQL_HANDLE_ENV  
    if (hType = SQL_HANDLE_ENV) 
      retv  = SQLSetEnvAttr(self.handle, SQL_ATTR_ODBC_VERSION, verId, SQL_IS_INTEGER);
    end 
  end

  return retv
! end allocateHandle 
! ----------------------------------------------------------------------------
  
OdbcHandleClType.freeHandle procedure() !,sqlReturn,proc

retv   sqlReturn,auto

  code
  
  retv = SqlFreeHandle(self.handleType, self.handle)    
  
  if (retv <> Sql_Success) 
    !self.getError(SQL_HANDLE_STMT, self.hStmt)
  else 
    self.handle = SQL_NO_HANDLE
  end 
  
  return retv
! end freeHandle 
! ----------------------------------------------------------------------------
  
OdbcHandleClType.getHandle procedure() !,SQLHANDLE

  code
  
  return self.handle
! end getHandle 
! ----------------------------------------------------------------------------  

! ------------------------------------------------------------
! unbinds the columns and the parameters for the statement 
! handle.  the statment handle at this stime could be reused 
! if needed/wanted
! ------------------------------------------------------------
OdbcStmtHandleClType.freeBindings  procedure() !,sqlReturn,proc,virtual

retv   sqlReturn,auto

  code 
   
  retv = SQLFreeStmt(self.handle, SQL_UNBIND)
  retv = SQLFreeStmt(self.handle, SQL_RESET_PARAMS)

  return retv
! ----------------------------------------------------------------------------

! ------------------------------------------------------------
! frees the statement handle 
! ------------------------------------------------------------
OdbcStmtHandleClType.freeHandle procedure() !,sqlReturn,proc,virtual

retv   sqlReturn,auto

  code 

  if (self.handle = SQL_NO_HANDLE) 
    return sql_Success
  end
  
  SQLFreeStmt(self.handle, SQL_CLOSE)
  self.freeBindings()
  
  retv = parent.freeHandle() 

  return retv  
! end freehhandle ----------------------------------------------  
  