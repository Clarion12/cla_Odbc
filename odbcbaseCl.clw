
  member()
  
  include('odbcBaseCl.inc'),once
  include('odbcSqlStrCl.inc'),once

  map 
    module('odbc32')
      SQLExecDirect(SQLHSTMT StatementHandle, odbcWideStr StatementText, SQLINTEGER TextLength),sqlReturn,pascal,raw,name('SQLExecDirectW')
      SQLCancel(SQLHSTMT StatementHandle),sqlReturn,pascal,proc
      SQLFetch(SQLHSTMT hs),sqlReturn,pascal
      SQLMoreResults(SQLHSTMT StatementHandle),sqlReturn,pascal
      SQLFreeStmt(SQLHSTMT StatementHandle, SQLUSMALLINT Option),sqlReturn,pascal
    end
  end

! ---------------------------------------------------------------------------
! Init 
! sets up the instance for use.  assigns the connection object input to the 
! data member.  allocates and init's the class to handle the sql statment or string. 
! ---------------------------------------------------------------------------  
odbcBaseClType.init procedure()   

retv     byte(level:benign)

  code 
  
  self.sqlStr &= new(sqlStrClType)
  if (self.sqlStr &= null)
    retv = level:notify
  else 
    self.sqlStr.init()
  end  
     
  self.errs &= new(ODBCErrorClType)
  if (self.errs &= null) 
    return sql_Error
  else 
    self.errs.Init()  
  end 

  return retv 
! end Init
! ----------------------------------------------------------------------

! ----------------------------------------------------------------------
! frees the memory used
! ----------------------------------------------------------------------
odbcBaseClType.kill procedure() !,virtual  

  code 

  if (~self.sqlStr &= null)
    self.sqlStr.kill() 
    self.sqlStr &= null
  end  
  
  if (~self.errs &= null) 
    dispose(self.errs)
    self.errs &= null
  end 

  return
! end kill
! ----------------------------------------------------------------------
 
! ----------------------------------------------------------------------
! default destructor, calls the kill method
! ---------------------------------------------------------------------- 
odbcBaseClType.destruct procedure()  !virtual

  code 

  self.kill()
  
  return
! end destructor
! ----------------------------------------------------------------------

! ----------------------------------------------------------------------
! unbinds any columns that are currently bound
! typical use is when a call that returns multiple result sets is in use.
! the first result set is processed and then call this to unbind the columns,
! then bind the caloums for the second result set, 
! repeat as needed
! ----------------------------------------------------------------------
odbcBaseClType.unBindColums procedure(SQLHSTMT hStmt)

retv sqlReturn
h  long

  code

  ! if there is a statment handle then unbind
  ! if not then nothing to do
  !if (self.conn.getHStmt() > 0) 
  !  self.conn.clearStmthandle()
  !end

  return
! end unBindColums
! ----------------------------------------------------------------------

! ----------------------------------------------------------------------
! virtual place holder
! use this function to format the fields or columns read prior to the display
! ----------------------------------------------------------------------
odbcBaseClType.formatRow procedure() !,virtual  

  code
  
  ! format queue elements for display in the derived object 
  
  return
! end formatRow 
! ----------------------------------------------------------------------

! -----------------------------------------------------------------------------
! Local worker function to assign the sql str (the actual sql statement) used in this 
! call to the class member
! -----------------------------------------------------------------------------  
odbcBaseClType.setSqlCommand procedure(*IdynStr s) ! sqlReturn,protected

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
odbcBaseClType.setSqlCommand procedure(string s) ! sqlReturn,protected

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
odbcBaseClType.nextResultSet procedure(SQLHSTMT hStmt) 

retv bool(false)
res  sqlReturn,auto

  code 
 
  res = SQLMoreResults(hStmt) 
  if (retv <> sql_success) and (retv <> SQL_SUCCESS_WITH_INFO)
    self.getError(hStmt)
    retv = false;
  else   
    retv = true
  end 
   
  return retv;
! ------------------------------------------------------------------------------

! ------------------------------------------------------------------------------
! reads the next result set by calling the fetch function
! ------------------------------------------------------------------------------
odbcBaseClType.readnextResult procedure(SQLHSTMT hStmt,  *queue q) !,sqlReturn,virtual

retv   sqlReturn,auto

  code

  retv = self.fetch(hStmt, q)
  if (retv <> sql_success) and (retv <> SQL_SUCCESS_WITH_INFO)
    self.getError(hStmt)
  end 

  return retv
! end  readnextResult
! ------------------------------------------------------------------------------

! ----------------------------------------------------------------------
! execute a query that does not return a result set and does not use any 
! parameters
! ----------------------------------------------------------------------
odbcBaseClType.executeStatement procedure(SQLHSTMT hStmt, *IDynStr sqlCode) !,sqlReturn,virtual

res     long,auto   ! used to avoid function call warnings
retv    sqlReturn,auto
wideStr CWideStr

  code 
  
  res = wideStr.Init(sqlCode.Cstr())

  retv = SQLExecDirect(hStmt, wideStr.GetWideStr(), SQL_NTS)
  if (retv <> sql_success) and (retv <> SQL_SUCCESS_WITH_INFO)
    self.getError(hStmt)
  end 

  return retv
! --------------------------------------------------------------------

! ------------------------------------------------------------------------------
! fetch with out a result set.  
! ------------------------------------------------------------------------------
odbcBaseClType.fetch procedure(SQLHSTMT hStmt) !sqlReturn,virtual

retv   sqlReturn,auto

  code 
  
  retv = SQLFetch(hStmt)
  if (retv <> sql_success) and (retv <> SQL_SUCCESS_WITH_INFO)
    self.getError(hStmt)
  end 
   
  return retv
! end fetch
! -----------------------------------------------------------------------------

! ------------------------------------------------------------------------------
! fetch
! reads the result set, one row at a time and places the data into the queue fields.
! Queue fields are already bound to the columns so all that is needed here is an add(q)
!
! Note, the queue fields must be bound before this method is called.
! ------------------------------------------------------------------------------
odbcBaseClType.fetch procedure(SQLHSTMT hStmt, *queue q) !sqlReturn,virtual

retv   sqlReturn,auto

  code 
  
  ! start loop and keep looping until an error or no_data is returned  
  
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
      ! format the queue elements for display, if needed, and add the element to the queue
      self.formatRow()
      add(q)
    else 
      ! dump the queue, something went wrong and 
      ! the code should not return a partial result set
      free(q)
      break    
    end  ! case
  end ! loop
  
  if (retv <> sql_success) and (retv <> SQL_SUCCESS_WITH_INFO) and (retv <> SQL_NO_DATA)
    self.getError(hStmt)
  end 

   
  return retv
! end fetch
! -----------------------------------------------------------------------------

odbcBaseClType.fetch procedure(SQLHSTMT hStmt, *queue q, *columnsClass cols) !sqlReturn,virtual

retv   sqlReturn,auto
allowNulls   bool,auto

  code 
  
  ! start loop and keep looping until an error or no_data is returned  
  
  allowNulls = cols.getallowNulls()
  
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
      if (allowNulls = true)
        cols.setDefaultNullValue(q)
      end
      ! format the queue elements for display, if needed, and add the element to the queue
      self.formatRow()
      add(q)
    else 
      ! dump the queue, something went wrong and 
      ! the code should not return a partial result set
      free(q)
      break    
    end  ! case
  end ! loop
  
  if (retv <> sql_success) and (retv <> SQL_SUCCESS_WITH_INFO) and (retv <> SQL_NO_DATA)
    self.getError(hStmt)
  end 
   
  return retv
! end fetch
! -----------------------------------------------------------------------------

! -----------------------------------------------------------------------------
! Binds the columns from the queue to the columns in the result set
! then calls fetch to read the result set
! -----------------------------------------------------------------------------
odbcBaseClType.fillResult procedure(SQLHSTMT hStmt, *columnsClass cols, *queue q, long setId = 1) !,sqlReturn,private

retv   sqlReturn,auto

  code 
 
  ! bind the columns just before the fetch, not needed for the execute query calls 
  ! so do it here, 
  retv = cols.bindColumns(Hstmt)

  ! if ok then go fetch the result
  if (retv <> sql_success) and (retv <> SQL_SUCCESS_WITH_INFO)
    self.getError(hStmt)
  else    
    retv = self.fetch(hStmt, q, cols)
    if (retv <> sql_success) and (retv <> SQL_SUCCESS_WITH_INFO)
      self.getError(hStmt)
    end
  end  

  return retv
! end fillResult
! -----------------------------------------------------------------------------

! -----------------------------------------------------------------------------
! call the error class to read the error information
! -----------------------------------------------------------------------------
odbcBaseClType.getError procedure(SQLHSTMT hStmt)  ! protected

retv   sqlReturn,auto

  code 
  
  self.errs.getError(SQL_HANDLE_STMT, hStmt)

  return
! end getError  
! -----------------------------------------------------------------------------
