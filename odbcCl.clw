

  member()
  
  include('odbcCl.inc'),once
  include('odbcSqlStrCl.inc'),once

  map 
    module('odbc32')
      SQLExecDirect(SQLHSTMT StatementHandle, odbcWideStr StatementText, SQLINTEGER TextLength),sqlReturn,pascal,raw,name('SQLExecDirectW')
      SQLExecute(SQLHSTMT StatementHandle),sqlReturn,pascal
      SQLCancel(SQLHSTMT StatementHandle),sqlReturn,pascal,proc
      SQLFetch(SQLHSTMT hs),sqlReturn,pascal
      SQLFetchScroll(SQLHSTMT StatementHandle, SQLSMALLINT FetchOrientation, SQLLEN FetchOffset),sqlReturn,pascal
      SQLPrepare(SQLHSTMT StatementHandle, *SQLCHAR StatementText, SQLINTEGER TextLength),sqlReturn,pascal,raw,name('SQLPrepareW')
      SQLCloseCursor(SQLHSTMT StatementHandle),sqlReturn,pascal
      SQLMoreResults(SQLHSTMT StatementHandle),sqlReturn,pascal
    end 
  end

! ---------------------------------------------------------------------------
! Init 
! sets upt he instance for use.  assigns the connection object input to the 
! data member.  allocates and init's the class to handle the sql statment or string. 
! ---------------------------------------------------------------------------  
odbcClType.init procedure(ODBCConnectionClType conn)   

retv     byte(level:benign)

  code 
  
  if (conn &= null) 
    return level:notify
  end
    
  self.conn &= conn
  self.sqlStr &= new(sqlStrClType)
  if (self.sqlStr &= null)
    retv = level:notify
  else 
    self.sqlStr.init()
  end  
     
  return retv 
! end Init
! ----------------------------------------------------------------------

odbcClType.kill procedure() !,virtual  

  code 

  if (~self.sqlStr &= null)
    self.sqlStr.kill() 
    self.sqlStr &= null
  end  
  
  self.conn &= null
  
  return
! end kill
! ----------------------------------------------------------------------
 
odbcClType.destruct procedure()  

  code 

  self.kill()
  
  return
! end destruct
! ----------------------------------------------------------------------

! virtual place holder
odbcClType.formatRow procedure() !,virtual  

  code
  
  ! format queue elements for display here
  
  return
! end formatRow 
! ----------------------------------------------------------------------

! -----------------------------------------------------------------------------
! Local worker function to assign the sql str (the actual sql statement) used in this 
! call to the class member
! -----------------------------------------------------------------------------  
odbcClType.setSqlCommand procedure(*IdynStr s) ! sqlReturn,protected

  code 
  
  ! make sure there is one
  if (s &= null) or (s.strLen() = 0)
    return sql_Error
  end 
  
  self.sqlStr.init(s)

  return sql_Success
! end setSqlCommand
! ----------------------------------------------------------------------  

! -----------------------------------------------------------------------------
! Local worker overloaded function to assign the sql str (the actual sql statement) used in this 
! call to the class member
! -----------------------------------------------------------------------------  
odbcClType.setSqlCommand procedure(string s) ! sqlReturn,protected

  code 
  
  ! make sure there is one
  if (len(clip(s)) = 0)
    return sql_Error
  end 
  
  self.sqlStr.init(s)

  return sql_Success
! end setSqlCommand
! ----------------------------------------------------------------------  

! ----------------------------------------------------------------------  
! checks for the next result set, if any, and moves to the next result set
! returns true if there is more and false if not
! ----------------------------------------------------------------------  
odbcClType.nextResultSet procedure() 

retv bool(false)
res  sqlReturn

  code 
 
  res = SQLMoreResults(self.conn.getHStmt()) 
  if (res = sql_success) 
    retv = true
  end 
   
  return retv;
! ------------------------------------------------------------------------------

! ------------------------------------------------------------------------------
! fetch
! reads the result set, one row at a time and places the data into the queue fields.
! Queue fields are already bound to the columns so all that is needed here is an add(q)
!
! Note, the queue fields must be bound before this method is called.
! ------------------------------------------------------------------------------
odbcClType.fetch procedure(*queue q) !sqlReturn,virtual

retv   sqlReturn
hStmt  SQLHSTMT

  code 
  
  ! start loop and keep looping until an error or no_data is returned  
  hStmt = self.conn.gethStmt()
  
  loop
    retv = SQLFetch(hStmt)
    case retv 
    of SQL_NO_DATA
      ! set back to success, no_data is expected (end of result set), 
      ! but caller is going to check for success
      retv = Sql_Success    
      break
    of Sql_Success
    orof Sql_Success_with_info
      add(q)
      self.formatRow()
    else 
      ! dump the queue, something went wrong and 
      ! the code should not return a partial result set
      free(q)
      break    
    end  ! case
  end ! loop
  
  if (retv = sql_success_with_info) 
    retv = sql_success
  end 
   
  return retv
! end fetch
! -----------------------------------------------------------------------------

! -----------------------------------------------------------------------------
! when using a scrollable cursor the cursor is closed by calling this function
! -----------------------------------------------------------------------------
odbcClType.closePaging procedure() !,sqlReturn

retv    sqlReturn,auto

  code 
  
  retv = SQLCloseCursor(self.conn.getHstmt())
  
  return retv
! -----------------------------------------------------------------------------

! -----------------------------------------------------------------------------
! function to fetch a page from the scrolling cursor,  
! sqlDir is an equate value, for this example will be next or previous 
! -----------------------------------------------------------------------------    
odbcClType.fetchPage  procedure(long sqlDir)

retv       sqlReturn,auto

  code 
  
  retv = SQLFetchScroll(self.conn.gethStmt(), sqlDir, 0)
  
  if (retv = sql_error)
    self.getError()
  end   
  
  return retv
! -----------------------------------------------------------------------------
  
! -----------------------------------------------------------------------------
! Binds the columns from the queue to the columns in the result set
! then calls fetch to read the result set
! -----------------------------------------------------------------------------
odbcClType.fillResult procedure(*columnsClass cols, *queue q) !,sqlReturn,private

retv   sqlReturn 

  code 
 
  ! bind the columns just before the fetch, not needed for the execute query calls 
  ! so do it here, 
  retv = cols.bindColumns(self.conn.getHstmt())

  ! if ok then go fetch the result
  if (retv = sql_success)
    retv = self.fetch(q)
  end  
  
  if (retv <> sql_Success)
    self.getError()
  end   

  return retv
! end fillResult
! -----------------------------------------------------------------------------

odbcClType.getError procedure() 

retv   sqlReturn
err    ODBCErrorClType

  code 
  
  retv = err.Init()
  err.getError(SQL_HANDLE_STMT, self.conn.getHstmt())
  
  return
! -----------------------------------------------------------------------------
    
! ------------------------------------------------------------------------------
! execQuery
! execute a query that returns a result set.  
! get a connection, prep the statement, execute the statement
! then fill the queue and close the connection when done
!
! this method does not accept the parameters class instance so use this one for queries that 
! do not have parameters.
! ------------------------------------------------------------------------------    
odbcClType.execQuery procedure(*IDynStr sqlCode, *columnsClass cols, *queue q) !,sqlReturn,virtual

retv    sqlReturn,auto

  code 
  
  if (self.setupQuery(sqlCode, cols) <> sql_Success)
    return sql_error
  end 
  
  retv = self.execQuery() 

  ! fill the queue
  if (retv = sql_Success)
    retv = self.fillResult(cols, q)
  end   
 
  return retv 
! end execQuery
! ----------------------------------------------------------------------

odbcClType.execQuery procedure(*sqlStrClType sqlCode, *columnsClass cols, *queue q) !,sqlReturn,virtual

retv    sqlReturn,auto

  code 
  
  if (self.setupQuery(sqlCode.sqlStr, cols) <> sql_Success)
    return sql_error
  end 
  
  !self.sqlStr.replaceFieldList(cols)
    
  retv = self.execQuery() 
  !self.getError()
  ! fill the queue
  if (retv = sql_Success)
    retv = self.fillResult(cols, q)
  end   
 
  return retv 
! end execQuery
! ----------------------------------------------------------------------


odbcClType.execQuery procedure(*IDynStr sqlCode) !,sqlReturn,virtual

res     long,auto
retv    sqlReturn,auto
wideStr CWideStr

  code 
  
  res = wideStr.Init(sqlCode.Cstr())
  retv = SQLExecDirect(self.conn.gethStmt(), wideStr.GetWideStr(), SQL_NTS)
   !sqlCode.Cstr()

  return retv
  
odbcClType.execQuery procedure(*IDynStr sqlCode, *columnsClass cols, *ParametersClass params, *queue q) !,sqlReturn,virtual

retv    sqlReturn,auto

  code 
  
  if (self.setupQuery(sqlCode, cols) <> sql_Success)
    return sql_error
  end 

  retv = self.execQuery(params) 
  
  ! fill the queue
  if (retv = sql_Success)
    retv = self.fillResult(cols, q)
  end   
  
  return retv 
! end execQuery
! ----------------------------------------------------------------------

odbcClType.execQuery procedure() !,sqlReturn,private

res     long,auto
retv    sqlReturn,auto
wideStr CWideStr

  code 
  
  res = wideStr.Init(self.sqlStr.cstr())
  !retv = SQLExecDirect(self.conn.gethStmt(), self.sqlStr.Cstr(), SQL_NTS)
  retv = SQLExecDirect(self.conn.gethStmt(), wideStr.getWideStr(), SQL_NTS)
  
  if (retV = sql_Success_with_info) 
    retv = sql_success
  end 
  if (retv <> sql_success)
    self.getError()
  end
    
  return retv
! -----------------------------------------------------------------------------
  
! -----------------------------------------------------------------------------
! bind the parameters and then call the execQuery/0 method
! -----------------------------------------------------------------------------  
odbcClType.execQuery procedure(*ParametersClass params) !,sqlReturn,private

retv    sqlReturn(sql_Success)

  code 
  
  ! if none then get out 
  if (params &= null)   
    return sql_error
  end   
    
  retv = params.bindParameters(self.conn.getHStmt(), self.sqlStr)
  
  if (retv = sql_Success)   
    retv = self.execQuery()
  end 
    
  return retv
! -----------------------------------------------------------------------------

odbcClType.prepQuery procedure(*IDynStr sqlCode) !,sqlReturn,virtual

retv   sqlReturn,auto

  code
 
  retv = SQLPrepare(self.conn.gethStmt(), sqlCode.cstr(), SQL_NTS)

  return retv
  
! -----------------------------------------------------------------------------
! execute a stored procedure, 
! function calls exec direct with the command text
! -----------------------------------------------------------------------------
odbcClType.execSp procedure() !private,sqlReturn

retv    sqlReturn,auto
wideStr CWideStr
retCount long

  code
 
!  retv = SQLPrepare(self.conn.gethStmt(), self.sqlStr.cstr(), SQL_NTS)
  !stop(retv)
  retCount = wideStr.Init(self.sqlStr.cstr())
  retv = SQLExecDirect(self.conn.gethStmt(), wideStr.GetWideStr(), self.sqlStr.strlen())
  
  if (retv <> Sql_Success) and (retv <> Sql_Success_with_info)
    self.getError()  
  end
  
  if (retv = Sql_Success_with_info)
    retv = sql_Success  
  end 
    
  return retv
! end execSp
! ---------------------------------------------------------------------------

! ------------------------------------------------------------------------------
! execSp
! call an stored procedure that does not return a value or a result set. 
! binds any parameters and calls execSp/0
! ------------------------------------------------------------------------------  
odbcClType.execSp procedure(string spName, *ParametersClass params) !,sqlReturn

retv    sqlReturn 

  code 
  
  if (self.setupSpCall(spName, params) <> sql_Success)
    return sql_Error
  end 
  
  retv = params.bindParameters(self.conn.gethStmt())
  
  if (retv = sql_Success) 
    retv = self.execSp()
  end  

  return retv 
! end execSp
! ----------------------------------------------------------------------

! ------------------------------------------------------------------------------
! execSp
! call an stored procedure that returns a result set, the 
! queue parameter is bound to the resutls, 
! sp does not expect any parameters
! ------------------------------------------------------------------------------  
odbcClType.execSp procedure(string spName, columnsClass cols, *queue q) !,sqlReturn

retv    sqlReturn 

  code 
  
  if (self.setupSpCall(spName) <> sql_Success)
    return sql_Error
  end 
  
  retv = self.execSp()
  if (retv = sql_Success)
    retv = self.fillResult(cols, q)
  end   
  
  return retv
! end execSp
! ----------------------------------------------------------------------

! -----------------------------------------------------------------------------
! execSp
! call an stored procedure that returns a result set, the 
! queue parameter is bound to the resutls, 
! binds any parameters and calls execSp/0 
! ------------------------------------------------------------------------------  
odbcClType.execSp procedure(string spName, columnsClass cols, *ParametersClass params, *queue q) !,sqlReturn

retv    sqlReturn 

  code 
  
  if (self.setupSpCall(spName, params) <> sql_Success)
    return sql_Error
  end 
    
  retv = params.bindParameters(self.conn.gethStmt())
    
  if (retv = sql_Success) 
    retv = self.execSp() 
    if (retv = sql_Success) 
      retv = self.fillResult(cols, q)
    end   
  end  

  return retv
! end execSp
! ----------------------------------------------------------------------
  
odbcClType.callScalar procedure(string spName, *ParametersClass params) 

retv    sqlReturn

  code 
  
  self.sqlStr.formatScalarCall(spName, params)
    
  retv = params.bindParameters(self.conn.gethStmt())
    
  if (retv = sql_Success) 
    retv = self.execSp() 
  end  

  return retv
! end execSp
! ----------------------------------------------------------------------

odbcClType.setupSpCall procedure(string spName) 

retv     sqlReturn,auto
params   &ParametersClass

  code 
  
  retv = self.setupSpCall(spName, params)
  
  return retv 
  
odbcClType.setupSpCall procedure(string spName, *ParametersClass params) ! sqlReturn,private

retv    sqlReturn 

  code 
  
  if (spName = '') 
    return sql_error
  end 

  if (params &= null) 
    self.sqlStr.formatSpCall(spName)
  else   
    self.sqlStr.formatSpCall(spName, params)
  end   
    
  return retv

odbcClType.setupQuery procedure(*IDynStr sqlCode, *columnsClass cols) !,sqlReturn,private

  code 
  
  if (self.setSqlCommand(sqlCode) <> sql_Success) 
    return sql_error
  end 
  
  self.sqlStr.replaceFieldList(cols)
 
  return sql_Success  