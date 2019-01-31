

  member()
  
  include('odbcCl.inc'),once
  include('odbcSqlStrCl.inc'),once

  map 
  end

! ---------------------------------------------------------------------------
! Init 
! sets up the instance for use.  assigns the connection object input to the 
! data member.  allocates and init's the class to handle the sql statment or string. 
! ---------------------------------------------------------------------------  
odbcExecType.init procedure()   

retv     byte(level:benign)

  code 
  
  retv = parent.init()
     
  return retv 
! end Init
! ----------------------------------------------------------------------

! ----------------------------------------------------------------------
! does any needed clean up and calls the parent kill
! ----------------------------------------------------------------------
odbcExecType.destruct procedure()  !virtual

  code 

  self.kill()
  
  return
! end destructor
! ----------------------------------------------------------------------

! ----------------------------------------------------------------------
! virtual place holder
! use this function to format the fields or columns read prior to the display
! ----------------------------------------------------------------------
odbcExecType.formatRow procedure() !,virtual  

  code
  
  ! format queue elements for display in the derived object 
  
  return
! end formatRow 
! ----------------------------------------------------------------------

! ----------------------------------------------------------------------
! execute a query that does not return a result set and does not use any 
! parameters
! ----------------------------------------------------------------------
odbcExecType.execQuery procedure(SQLHSTMT hStmt, *IDynStr sqlCode) !,sqlReturn,virtual

res     long,auto   ! used t oavoid function call warnings
retv    sqlReturn,auto

  code 
  
  retv = self.executeStatement(hStmt, sqlCode)

  return retv
! --------------------------------------------------------------------

! ------------------------------------------------------------------------------
! execQuery
! execute a query that returns a result set.  
! prep the statement, execute the statement
! then fill the queue or buffers and close the connection when done
!
! this method does not accept the parameters class instance so use this one for queries that 
! do not have parameters.
! ------------------------------------------------------------------------------    
odbcExecType.execQuery procedure(SQLHSTMT hStmt, *IDynStr sqlCode, *columnsClass cols, *queue q) !,sqlReturn,virtual

retv    sqlReturn,auto

  code 
  
  retv = self.execQuery(hStmt, sqlCode)

  ! fill the queue
  if (retv = sql_Success) or (retv = SQL_SUCCESS_WITH_INFO)
    retv = self.fillResult(hStmt, cols, q)
  end   
 
  return retv 
! end execQuery
! ----------------------------------------------------------------------

! ----------------------------------------------------------------------
! execute a query that returns a result set and expects parameters
! do the set up and bind the parameters and execute, then fill the 
! result set
! ----------------------------------------------------------------------
odbcExecType.execQuery procedure(SQLHSTMT hStmt, *IDynStr sqlCode, *columnsClass cols, *ParametersClass params, *queue q) !,sqlReturn,virtual

retv    sqlReturn,auto

  code 

  retv = params.bindParameters(hStmt)

  retv = self.execQuery(hStmt, sqlCode)
  
  ! fill the queue
  if (retv = sql_Success) or (retv = SQL_SUCCESS_WITH_INFO)
    retv = self.fillResult(hStmt, cols, q)
  end   
  
  return retv 
! end execQuery
! ----------------------------------------------------------------------

! ------------------------------------------------------------------------------
! execQuery
! execute a query that does not return a result set.  
! but has at least one output parameter
! ------------------------------------------------------------------------------      
odbcExecType.execQueryOut procedure(SQLHSTMT hStmt, *IDynStr sqlCode, *ParametersClass params) !,sqlReturn,virtual

retv    sqlReturn,auto

  code 
  
  retv = self.execQuery(hStmt, sqlCode, params)
  
  return retv 
! end execQuery
! ----------------------------------------------------------------------

! ------------------------------------------------------------------------------
! execQuery
! execute a query that does not return a result set.  
! but has input parameters and may have output parameters.
! ------------------------------------------------------------------------------      
odbcExecType.execQuery  procedure(SQLHSTMT hStmt, *IDynStr sqlCode, *ParametersClass params) !,sqlReturn,virtual

retv  sqlReturn,auto 

  code

  retv = params.bindParameters(hStmt)
  if (retv = sql_Success) or (retv = SQL_SUCCESS_WITH_INFO)
    retv = self.execQuery(hStmt, sqlCode)
  end  

  return retv
! end execQuery
! ----------------------------------------------------------------------