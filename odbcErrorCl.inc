
!ABCIncludeFile

  OMIT('_EndOfInclude_',_ODBCErrorClPresent_)
_ODBCErrorClPresent_   EQUATE(1)

  include('odbcTypes.inc'),once
  include('odbcConn.inc'),once
  include ('svcom.inc'),once

    
eLengthSqlState    equate(6)
eLengthErrText     equate(2000)
  
OdbcErrorGroup group,type
sqlState         cstring(eLengthSqlState)
NativeErrorPtr   long
MessageText      cstring(eLengthErrText)
textLengthPtr    SQLSMALLINT
               end 

OdbcErrorQueue queue(OdbcErrorGroup)
               end 
               
ODBCErrorClType class,type,module('odbcErrorCl.clw'),link('odbcErrorCl.clw',_ABCLinkMode_),dll(_ABCDllMode_)
errorMsgQ         &OdbcErrorQueue,private
displayError      bool,private    ! display an error using the message function when it happens
errorCount        long,private    ! number of error records retruned by the most recent error for the handle type

! ----------------------------------------------------------------
! setup and clean up
! ----------------------------------------------------------------
init              procedure(),sqlReturn
kill              procedure(),virtual
destruct          procedure(),virtual
!
! ----------------------------------------------------------------
! toggles the flag t oshow the error or notshow the error
! true the error is displayed when it happens
! ----------------------------------------------------------------
setDisplayError   procedure(bool onOff)

getConnError      procedure(ODBCConnectionClType conn),sqlReturn,proc 
getEnvError       procedure(ODBCConnectionClType conn),sqlReturn,proc
getDataBaseError  procedure(ODBCConnectionClType conn),sqlReturn,proc

getError          procedure(SQLSMALLINT HandleType, SQLHANDLE Handle),sqlReturn,proc

! ----------------------------------------------------------------------  
! calls the SQLGetDiagField function to get the number of messages 
! they maybe error messages or just information messages
! called internally by the getError function
! this call can return other information but they are not currently used
!
! the call can return the number of rows affected by an insert, update or delete action
! and some other types of information
! ----------------------------------------------------------------------    
getDiagRecCount   procedure(SQLSMALLINT HandleType, SQLHANDLE Handle),long,private,proc

! ----------------------------------------------------------------------
! displays an erro on the screen using a simple message call
! shows all the errors and information messages from the call for the handle type
! ----------------------------------------------------------------------
showError         procedure()

! ----------------------------------------------------------------------
! free the queue of error messages
! ----------------------------------------------------------------------
freeErrorMsgQ     procedure(),private

makeObjects       procedure(),sqlReturn,private
destroyObjects    procedure(),private
                end ! class 


  _EndOfInclude_
