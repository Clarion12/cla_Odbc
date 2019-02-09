!ABCIncludeFile

  OMIT('_EndOfInclude_',_ODBCTransClassPresent_)
_ODBCTransClassPresent_   EQUATE(1)

  include('odbcTypes.inc'),once

odbcTransactionClType class,type,module('odbcTransCl.clw'),link('odbcTransCl.clw',_ABCLinkMode_),dll(_ABCDllMode_)

! ------------------------------------------------------
! the connection handle this instance will use for 
! the various calls.  
! ------------------------------------------------------
hDbc     SQLHDBC,private

! -------------------------------------------------------------------
! the isolation level set by the user for 
! a transaction. used to reset to the default when a transaction ends
! -------------------------------------------------------------------
currentIsolationLvl SQLINTEGER

! ------------------------------------------------------------------
! the default isolation leve that will be used by the transactions
! set this to what ever is needed/wnated 
! but unless there is very specific reasons to change the default of 
! read committed should be used.
! ------------------------------------------------------------------
defaultIsolationLvl SQLINTEGER,private 

! set up and tear down
init             procedure(SQLHDBC hanlde),byte,proc
kill             procedure(),virtual
destruct         procedure(),virtual

! ----------------------------------------------------------------------
! sets the default value for the isolation level
! read committed is the normal default and is set in the constructor
! there may be use cases where some other default value is needed/wanted.
! However, if altering the default be sure you understand the implications 
! of using the other types of isolation levels.
! ----------------------------------------------------------------------
setDefaultIsolationLevel procedure(SQLINTEGER level)

! ----------------------------------------------------------------------
! sets the isolation level for the hDbc input.  the connection must not 
! have any open transactions when this is called.  the connection can be open 
! or closed.  as a general rule the connection would be open but that is 
! not required.  
! Note, if there are open transactions when this is called bad things will happen 
! ----------------------------------------------------------------------
setIsolationLevel procedure(long level),sqlReturn,protected 

! ----------------------------------------------------------------------
! sets the isolation level for the hDbc input to the level indicated 
! by the function name.  typically the default value of read committed is all that 
! will be used, there are use case for serializable transaction and in some 
! other instances the two types are used.
! ----------------------------------------------------------------------
setIsolationSerializable procedure(),sqlReturn 
setIsolationReadUncommitted procedure(),sqlReturn
setIsolationReadCommitted procedure(),sqlReturn
setIsolationRepeatabelread procedure(),sqlReturn

! ----------------------------------------------------------------------
! begins a transaction for the connection handle input.  
! this actually turns off auto-commit mode.
! ----------------------------------------------------------------------
beginTrans procedure(),sqlReturn

! ----------------------------------------------------------------------
! commits a transaction for the connection handle input.  
! ----------------------------------------------------------------------
Commit procedure(),sqlReturn

! ----------------------------------------------------------------------
! rolls back a transaction for the connection handle input.  
! ----------------------------------------------------------------------
Rollback procedure(),sqlReturn

! ----------------------------------------------------------------------
! ends a transaction for the connection handle input.  
! called from the commit or rollback functions.
! ----------------------------------------------------------------------
EndTrans procedure(long committRollBack),sqlReturn,private

            end

  _EndOfInclude_
