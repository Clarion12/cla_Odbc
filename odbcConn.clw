
  member()
  
  include('odbcConn.inc'),once 
  include'odbcTypes.inc'),once
  include('svcom.inc'),once

  map 
    module('odbc32')
      SQLConnect(SQLHDBC ConnectionHandle, *SQLCHAR ServerName, SQLSMALLINT NameLength1, long UserName, SQLSMALLINT NameLength2, long Authentication, SQLSMALLINT NameLength3),sqlReturn,pascal,raw
      SQLDriverConnect(SQLHDBC ConnectionHandle, SQLHWND WindowHandle, long InConnectionString, SQLSMALLINT StringLength1, long  OutConnectionString, SQLSMALLINT BufferLength, *SQLSMALLINT StringLength2Ptr, SQLUSMALLINT DriverCompletion),sqlReturn,pascal,raw,Name('SQLDriverConnectW')
      SQLDisconnect(SQLHDBC ConnectionHandle),sqlReturn,pascal   
      SQLGetInfo(long hDbc, long attrib, *cstring valuePtr, long buffLength, *long strLenPtr),long,pascal,raw
      SQLSetEnvAttr(SQLHENV EnvironmentHandle, SQLINTEGER Attribute,  SQLPOINTER Value, SQLINTEGER StringLength),sqlReturn,pascal
    end 
  end

eNoWindow  equate(0)

ODBCConnectionClType.init procedure() !,sqlReturn

retv    sqlReturn,auto

  code 

  self.hEnv &= new(OdbcHandleClType) 
  if (self.hEnv &= null) 
    return sql_Error
  end
      
  self.hDbc &= new(OdbcHandleClType)
  if (self.hDbc &= null) 
    return sql_Error
  end   

  self.hStmt &= new(OdbcStmtHandleClType)
  if (self.hStmt &= null) 
    return sql_Error
  end   
  
  retv = self.hEnv.allocateHandle(SQL_HANDLE_ENV, Sql_Null_Handle)
  
  return retv
! end init
! -------------------------------------------------------------------------

ODBCConnectionClType.init procedure(*IConnString connString) !,sqlReturn

retv    sqlReturn,auto

  code 

  if (connString &= null) 
    return sql_Error
  end    

  self.connStr &= connString

  retv = self.Init()
    
  return retv
! end init
! -------------------------------------------------------------------------
    
ODBCConnectionClType.kill procedure()

  code 
  
  if (~self.hDbc &= null)
    dispose(self.hdbc)
    self.hDbc &= null
  end 
  
  if (~self.hEnv &= null)
    dispose(self.hEnv)
    self.hEnv &= null
  end 
  
  self.connStr &= null
  if (~self.hStmt &= null)
    dispose(self.hStmt)
    self.hStmt &= null
  end
  
  return 
! end kill 
! -------------------------------------------------------------------------  

ODBCConnectionClType.gethEnv procedure() !,SQLHEnv

  code 
  return self.hEnv.gethandle()
  
ODBCConnectionClType.gethDbc procedure() !,SQLHDBC

  code 
  return self.hDbc.getHandle()
 
ODBCConnectionClType.gethStmt procedure() !,SQLHStmt

  code 
  return self.hStmt.gethandle()
! ---------------------------------------------------------------------------

ODBCConnectionClType.setOdbcVersion procedure() 

err             ODBCErrorClType
retv            sqlReturn
SQL_OV_ODBC3_80 long(380)

  code

  retv  = SQLSetEnvAttr(self.gethEnv(), SQL_ATTR_ODBC_VERSION, SQL_OV_ODBC3_80, SQL_IS_INTEGER);
  if (retv <> Sql_Success) 
    err.getError(SQL_HANDLE_ENV, self.gethEnv())
  end

  return retv
! --------------------------------------------------------------------------

ODBCConnectionClType.connect procedure(bool statement = withStatement)

retv       sqlReturn,auto
outConnStr cstring(2000)
outLength  sqlsmallint
holder     cstring(2000)
wideStr    Cwidestr
outWdireStr Cwidestr

  code 

  if (self.hdbc.getHandle() <= 0)
    retv = self.hDbc.allocateHandle(SQL_HANDLE_DBC, self.hEnv.getHandle())
  else
    retv = sql_Success
  end

  if (retv = sql_Success) or (retv = sql_success_with_info)
    holder = self.connStr.ConnectionString()
    wideStr.init(holder)
    outConnStr = all(' ')
    outWdireStr.Init(outConnStr)
    retv = SQLDriverConnect(self.hDbc.getHandle(), eNoWindow, widestr.GetWideStr(), SQL_NTS, outWdireStr.GetWideStr(), size(outConnStr) - 1, outLength, SQL_DRIVER_NOPROMPT)
  end   

  ! check for with info, always returns with info about the connection
  if (retv <> sql_Success) and (retv <> Sql_Success_With_Info)
    !self.getDatabaseError()
  else 
    if (statement = true) 
      retv = self.hStmt.AllocateHandle(SQL_HANDLE_STMT, self.hDbc.getHandle())
      if (retv <> sql_Success) and (retv <> Sql_Success_With_Info)
        !self.getDatabaseError()
      end   
    end   
  end
 
  if (retv = Sql_Success_With_Info) 
   ! reset for the caller
    retv = sql_Success
  end   
  
  return retv 
! end connect 
! ----------------------------------------------------------------------

ODBCConnectionClType.Disconnect procedure(bool statement = withStatement)

retv      sqlReturn
h         SQLHDBC,auto

  code 

  if (statement = true) 
    self.hStmt.freeHandle()
  end

  h = self.hDbc.getHandle()
  if (h > 0)
    retv = SQLDisconnect(h)
    if (retv <> sql_Success) and (retv <> sql_Success_With_Info)
      self.getDatabaseError()
    else
      ! if with info reset for the caller
      retv = sql_Success  
    end
  end 

  return retv
! end Disconnect 
! ----------------------------------------------------------------------
  
ODBCConnectionClType.getDatabaseError procedure() !,virtual,protected  

err    ODBCErrorClType

  code 
  
  err.getDatabaseError(self)
  
  return 
  
ODBCConnectionClType.getError procedure(SQLSMALLINT HandleType, SQLHANDLE Handle)  

err    ODBCErrorClType

  code 
  
  err.getError(handleType, handle)
      
  return 
! end getError 
! ----------------------------------------------------------------------