
  member()
  
  include('odbcConnStrCl.inc'),once 

  map 
  end

eTrustedConnTextOn   equate('Trusted_Connection=yes;')
eTrustedConnTextOff  equate('Trusted_Connection=no;')

eDriverLabel         equate('Driver={{')
eServerLabel         equate('Server=')
eDbLabel             equate('Database=')

eConnDelimit      equate(';')

! -----------------------------------------------------------------
! init/0
! allocate the dyn string for the connection class instance
! -----------------------------------------------------------------
baseConnStrClType.init procedure(long db) !,*IConnString

retv     byte(level:benign)
ms       &MSConnStrClType
postGre  &PgConnStrClType

  code 
  
  case db 
  of dbVendor:MS
    ms &= new(MSConnStrClType)
    ms.connStr &= newDynStr()
    ! if it fails return null to indicate an error
    if (ms.connStr &= null)
      self.IConnStr &= null
    else 
      self.IConnStr &= ms.ISqlConnStr
    end  ! if (ms.connStr &= null)
    self.trustedConn = true
  of dbVendor:PostGre
    postGre &= new(PgConnStrClType)
    postGre.connStr &= newDynStr()
    if (postgre.connStr &= null) 
      self.IConnStr &= null
    else
      self.IConnStr &= postGre.IPgConnStr
    end ! if (postgre.connStr &= null) 
   ! add other vendors, example
   !of dbVendor:Oracle    
   end ! case db

  return self.IConnStr
! end init 
! -------------------------------------------------------------------------------------

! -------------------------------------------------------------------------------------
! Init 
! set up the instance based on the db input and uses the three string input to prime 
! the fields.  
! -------------------------------------------------------------------------------------
baseConnStrClType.init procedure(long db, string driverName, string srvName, string dbName) !,*IConnString

  code 
  
  ! this param must be present and in range
  if (db <= 0) or (db >= dbVendor:lastVendor)
    return null
  end 
     
  if (~self.init(db) &= null)
    self.IConnStr.setDriverName(driverName)
    self.IConnStr.setSrvName(srvName)
    self.IConnStr.setDbName(dbName)
    ! assume trusted security is in use for this method, over load if needed or 
    ! call the setter after this call  
    self.IConnStr.setTrustedConn(true)
  end  
  
  return self.IConnStr
! end init 
! -----------------------------------------------------------------
  
! -----------------------------------------------------------------
! kill
! dispose the dyn string 
! -----------------------------------------------------------------
baseConnStrClType.kill procedure()

  code 
  
  disposeDynStr(self.connStr)
  self.connStr &= null
  
  return
! end kill
! -----------------------------------------------------------------

MSConnStrClType.ISqlConnStr.kill procedure()

  code 

  self.kill()
  dispose(Self)
   
  return 
! end kill
! ------------------------------------------------------------------------------

MSConnStrClType.ISqlConnStr.ConnectionString procedure() !,*cstring

  code 
  
  ! clear it 
  self.connStr.Kill()
  ! and then build it
  
  ! add the defaults that are always used
  self.connStr.cat(self.driverName & self.SrvName &  self.dbName)

  if (self.TrustedConn = true)
    self.connStr.cat(eTrustedConnTextOn & eConnDelimit)
  else 
    self.connStr.cat(eTrustedConnTextOff)
  end  
     
  return self.connStr.cstr()
! end ConnectionString
! ------------------------------------------------------------------------------

! -----------------------------------------------------------------
! Setters for the instance
! -----------------------------------------------------------------
MSConnStrClType.ISqlConnStr.setDriverName procedure(string driverName)

  code 
  
  self.driverName = eDriverLabel & clip(driverName) & '}'  & eConnDelimit
  
  return
! end setDbName
! ------------------------------------------------------------------------------
  
MSConnStrClType.ISqlConnStr.setDbName procedure(string dbname)

  code 
  
  self.dbName = eDbLabel & clip(dbName) & eConnDelimit
  
  return
! end setDriverName
! ------------------------------------------------------------------------------
  
MSConnStrClType.ISqlConnStr.setSrvName procedure(string srvName)

  code 
  
  self.srvName = eServerLabel & clip(srvName) & eConnDelimit
  
  return
! end setServerName
! ------------------------------------------------------------------------------
  
MSConnStrClType.ISqlConnStr.setUserName procedure(string  user)

  code 
  
  self.userName = 'User ID=' & clip(user) & eConnDelimit
  
  return
! end setUserName
! ------------------------------------------------------------------------------
  
MSConnStrClType.ISqlConnStr.setPassword procedure(string pw)

  code 
  
  self.password = 'Password=' & clip(pw) & eConnDelimit
  
  return
! end setpassword
! ------------------------------------------------------------------------------
  
MSConnStrClType.ISqlConnStr.setTrustedConn procedure(bool onOff)

  code 

  self.trustedConn = onOff  
  
  return
! end setTrustedConn
! ------------------------------------------------------------------------------

MSConnStrClType.ISqlConnStr.setUseMars procedure(bool onOff)

  code 
  
  self.useMars = onOff
  
  return 
! end setUseMars
! ------------------------------------------------------------------------------
  
MSConnStrClType.ISqlConnStr.setPortNumber procedure(string portNumber)

  code
  
  self.portNumber = 'Port=' & clip(portNumber) & eConnDelimit
  
  return
! end setPortNumber
! ------------------------------------------------------------------------------


! end field setters for MS SQL Server 
! ------------------------------------------------------------------------------

! begin PostgreSql connection string interface mathods  

pgConnStrClType.IPgConnStr.kill procedure()

  code 

  self.kill()
  dispose(Self)
   
  return 
! end kill 
! ------------------------------------------------------------------------------
    
pgConnStrClType.IPgConnStr.ConnectionString procedure() !,*cstring

  code 
  
  ! clear it 
  self.connStr.Kill()
  ! and then build it
  ! example connection string, DNS less
  !Driver={PostgreSQL};Server=IP address;Port=5432;Database=myDataBase;Uid=myUsername;Pwd=myPassword                   
  self.connStr.cat(self.driverName & self.SrvName & self.portNumber &  self.dbName & self.userName & self.passWord)

  return self.connStr.cstr()
! end ConnectionString
! ------------------------------------------------------------------------------
  
! -----------------------------------------------------------------
! Setters for the instance
! -----------------------------------------------------------------
pgConnStrClType.IpgConnStr.setDriverName procedure(string driverName)

  code 
  
  self.driverName = eDriverLabel & clip(driverName) &'}' & eConnDelimit 
  
  return
! end setDrivername
! ------------------------------------------------------------------------------
  
pgConnStrClType.IpgConnStr.setDbName procedure(string dbname)

  code 
  
  self.dbName = eDbLabel & clip(dbName) & eConnDelimit
  
  return
! end setDbName
! ------------------------------------------------------------------------------
  
pgConnStrClType.IpgConnStr.setSrvName procedure(string srvName)

  code 
  
  self.srvName = eServerLabel & clip(srvName) & eConnDelimit
  
  return
! end setSrvName
! ------------------------------------------------------------------------------
  
pgConnStrClType.IpgConnStr.setUserName procedure(string  user)

  code 
  
  self.userName = 'Uid='& clip(user) & eConnDelimit
  
  return
! end setUserName
! ------------------------------------------------------------------------------
  
pgConnStrClType.IpgConnStr.setPassword procedure(string pw)

  code 
  
  self.password = 'Pwd='& clip(pw) & eConnDelimit
  
  return
! end setPassword
! ------------------------------------------------------------------------------

pgConnStrClType.IpgConnStr.setPortNumber procedure(string portNumber)

  code
  
  self.portNumber = 'Port=' & clip(portNumber) & eConnDelimit
  
  return
! end setPortNumber
! ------------------------------------------------------------------------------
    
pgConnStrClType.IpgConnStr.setTrustedConn procedure(bool onOff)

  code 

  self.trustedConn = onOff  
  
  return
! end setTrustedConn
! ------------------------------------------------------------------------------

! end field setters   


! begin Oracle connection string interface methods  

oraConnStrClType.IoraConnStr.kill procedure()

  code 

  self.kill()
  dispose(Self)
   
  return 
! end kill 
! ------------------------------------------------------------------------------
    
oraConnStrClType.IoraConnStr.ConnectionString procedure() !,*cstring

  code 
  
  ! clear it 
  self.connStr.Kill()
  ! and then build it
  ! example connection string, DNS less
  !Driver=(Oracle in XEClient);dbq=localhost/XE;Uid=tom;Pwd=tomtom1;
  self.connStr.cat(self.driverName & self.SrvName & self.userName & self.passWord)

  return self.connStr.cstr()
! end ConnectionString
! ------------------------------------------------------------------------------
  
! -----------------------------------------------------------------
! Setters for the instance
! -----------------------------------------------------------------
oraConnStrClType.IoraConnStr.setDriverName procedure(string driverName)

  code 
  
  self.driverName = eDriverLabel & clip(driverName) &'}' & eConnDelimit 
  
  return
! end setDrivername
! ------------------------------------------------------------------------------
  
oraConnStrClType.IoraConnStr.setDbName procedure(string dbname)

  code 
  
  self.dbName = eDbLabel & clip(dbName) & eConnDelimit
  
  return
! end setDbName
! ------------------------------------------------------------------------------
  
oraConnStrClType.IoraConnStr.setSrvName procedure(string srvName)

  code 
  
  self.srvName = 'dbq=' & clip(srvName) & '/XE' & eConnDelimit
  
  return
! end setSrvName
! ------------------------------------------------------------------------------
  
oraConnStrClType.IoraConnStr.setUserName procedure(string  user)

  code 
  
  self.userName = 'Uid='& clip(user) & eConnDelimit
  
  return
! end setUserName
! ------------------------------------------------------------------------------
  
oraConnStrClType.IoraConnStr.setPassword procedure(string pw)

  code 
  
  self.password = 'Pwd='& clip(pw) & eConnDelimit
  
  return
! end setPassword
! ------------------------------------------------------------------------------

oraConnStrClType.IoraConnStr.setPortNumber procedure(string portNumber)

  code
  
  self.portNumber = 'Port=' & clip(portNumber) & eConnDelimit
  
  return
! end setPortNumber
! ------------------------------------------------------------------------------
    
oraConnStrClType.IoraConnStr.setTrustedConn procedure(bool onOff)

  code 

  self.trustedConn = onOff  
  
  return
! end setTrustedConn
! ------------------------------------------------------------------------------

! end field setters   