   PROGRAM


NetTalk:TemplateVersion equate('14.23')
ActivateNetTalk   EQUATE(1)
  include('NetAll.inc'),once
  include('NetMap.inc'),once
  include('NetTalk.inc'),once
  include('NetSimp.inc'),once
  include('NetFtp.inc'),once
  include('NetHttp.inc'),once
  include('NetWww.inc'),once
  include('NetSync.inc'),once
  include('NetWeb.inc'),once
  include('NetWebSessions.inc'),once
  include('NetWebSocketClient.inc'),once
  include('NetWebSocketServer.inc'),once
  include('NetWebM.inc'),once
  include('NetWSDL.inc'),once
  include('NetEmail.inc'),once
  include('NetFile.inc'),once
  include('NetWebSms.inc'),once
  Include('NetOauth.inc'),once
  Include('NetLDAP.inc'),once
  Include('NetMaps.inc'),once
  Include('NetDrive.inc'),once
  Include('NetSms.inc'),once
  Include('NetDns.inc'),once
StringTheory:TemplateVersion equate('3.69')
jFiles:TemplateVersion equate('3.08')
Reflection:TemplateVersion equate('1.27')
ResizeAndSplit:TemplateVersion equate('5.10')
FM3:Version           equate('5.62')       !Deprecated - but exists for backward compatibility
FM3:TemplateVersion   equate('5.62')
GPFRep:Version equate ('2.37')           !Deprecated - but exists for backward compatibility
GPFReporter:TemplateVersion equate ('2.37')
HyperActive:TemplateVersion equate('2.34')
CapesoftMessageBox:TemplateVersion equate('2.49')
WinEvent:TemplateVersion      equate('5.41')

   INCLUDE('ABERROR.INC'),ONCE
   INCLUDE('ABFILE.INC'),ONCE
   INCLUDE('ABUTIL.INC'),ONCE
   INCLUDE('ERRORS.CLW'),ONCE
   INCLUDE('KEYCODES.CLW'),ONCE
   INCLUDE('ABFUZZY.INC'),ONCE
  include('cwsynchc.inc'),once  ! added by NetTalk
  include('StringTheory.Inc'),ONCE
   include('jFiles.inc'),ONCE
  include('Reflection.Inc'),ONCE
  include('ResizeAndSplit.Inc'),ONCE
  include('csGPF.Inc'),ONCE
   include('Hyper.Inc'),ONCE
    Include('dwrTrace.inc'),ONCE
    Include('WinEvent.Inc'),Once

   MAP
     MODULE('HTMLPAGESCRAPER_BC.CLW')
DctInit     PROCEDURE                                      ! Initializes the dictionary definition module
DctKill     PROCEDURE                                      ! Kills the dictionary definition module
     END
!--- Application Global and Exported Procedure Definitions --------------------------------------------
     MODULE('HTMLPAGESCRAPER001.CLW')
Main                   PROCEDURE   !
     END
     MODULE('HTMLPAGESCRAPER003.CLW')
RuntimeFileManager     PROCEDURE   !
     END
     MODULE('HTMLPAGESCRAPER005.CLW')
ds_Halt                PROCEDURE(UNSIGNED Level=0,<STRING HaltText>)   !
     END
     MODULE('HTMLPAGESCRAPER006.CLW')
ds_Stop                PROCEDURE(<string StopText>)   !
     END
     MODULE('HTMLPAGESCRAPER007.CLW')
ds_Message             FUNCTION(STRING MessageTxt,<STRING HeadingTxt>,<STRING IconSent>,<STRING ButtonsPar>,UNSIGNED Defaults=0,BOOL StylePar=FALSE),UNSIGNED,PROC   !
     END
       include('FM3map.clw')
       MyOKToEndSessionHandler(long pLogoff),long,pascal
       MyEndSessionHandler(long pLogoff),pascal
   END

  include('StringTheory.Inc'),ONCE
Glo:st               StringTheory
SilentRunning        BYTE(0)                               ! Set true when application is running in 'silent mode'

!region File Declaration
Projects             FILE,DRIVER('TOPSPEED'),NAME('Projects.tps'),PRE(Pro),CREATE,BINDABLE,THREAD !                     
PKProGuidKey             KEY(Pro:Guid),NOCASE,PRIMARY      !                     
ProDescriptionKey        KEY(Pro:Description),DUP,NOCASE   !                     
ProMainUrlKey            KEY(Pro:MainUrl),DUP,NOCASE       !                     
Record                   RECORD,PRE()
Guid                        STRING(16)                     !                     
Description                 STRING(50)                     !                     
MainUrl                     STRING(100)                    !                     
SaveTo                      STRING(256)                    !                     
                         END
                     END                       

!endregion

ds_VersionModifier  long
ds_FMInited byte
gTopSpeedFile File,driver('TopSpeed',''),pre(__gtps)
record         record
a                byte
               end
             end



! ----- ThisGPF --------------------------------------------------------------------------
ThisGPF              Class(GPFReporterClass)
    ! derived method declarations
Construct              PROCEDURE ()
Destruct               PROCEDURE ()
_GetSymbol             PROCEDURE (ulong pAddress,byte pStackTrace=1),string ,VIRTUAL
_LookupExceptionCode   PROCEDURE (ulong p_ExceptionCode),string ,VIRTUAL
_VectoredExceptionHandler_ PROCEDURE (ulong p_e),long ,VIRTUAL
_StackDetails          PROCEDURE (ulong p_e,byte p_details,ulong p_hProcess),string ,VIRTUAL
_LocateDebugSymbols    PROCEDURE (long phModule),byte ,VIRTUAL
_ReadBlockFromFile     PROCEDURE (ulong pOffset,long pReadBytes,*string pFileBlock,*string pFileName) ,VIRTUAL
_GetModuleName         PROCEDURE (long phModule),string ,VIRTUAL
_GetModuleHandle       PROCEDURE (ulong pAddress),long ,VIRTUAL
LookupAddress          PROCEDURE () ,VIRTUAL
Initialize             PROCEDURE () ,VIRTUAL
FilterExceptions       PROCEDURE (ulong pException) ,VIRTUAL
_EncodeEmail           PROCEDURE (string pEmailText),string ,VIRTUAL
_FindFirstBreak        PROCEDURE (string pText,long pMaxLen),long ,VIRTUAL
ExtraReportText        PROCEDURE () ,VIRTUAL
_GetDLLVersion         PROCEDURE (string pDLLName),string ,VIRTUAL
_SetFileNames          PROCEDURE () ,VIRTUAL
DeleteDumpFile         PROCEDURE () ,VIRTUAL
_InitReportText        PROCEDURE () ,VIRTUAL
_ExecuteCommands       PROCEDURE () ,VIRTUAL
_StackDump             PROCEDURE (long pStart,long pEnd,long pStackLevel),string ,VIRTUAL
_FindLinePosition      PROCEDURE (string pText,long pLineNumber),long ,VIRTUAL
_DebugLog              PROCEDURE (string pDebugData,byte pFirstLine=0) ,VIRTUAL
_FormatLineInfo        PROCEDURE (long pLineNumber,string pProcName,string pSourceName,string pModuleName,byte pStackTrace,byte pNoProcFound,byte pExactAddress,byte pNoLineNumber),string ,VIRTUAL
_GetAssert             PROCEDURE (long pSP,long pBP),long ,VIRTUAL
_GetOtherMessage       PROCEDURE (long pSP,long pBP),long ,VIRTUAL
_RestartProgram        PROCEDURE () ,VIRTUAL
                     End  ! ThisGPF
! ----- end ThisGPF -----------------------------------------------------------------------
         include('MessageBox.inc'),once
ThisMessageBoxGlobal class(csThreadSafeMessageClass)
                     end

ThisMessageBox       class(csEnhancedMessageClass) ,thread
AssignGlobalClass        procedure    ,virtual
Init                     procedure  (long UseABCClasses=0,long UseDefaultFile=0)  ,virtual
PrimeLog                 procedure  (<string ExtraDetails>)  ,virtual
                      end
StartSearchEXEName          long(1)
EndSearchEXEName            long(0)
WE::MustClose       long
WE::CantCloseNow    long
  compile ('****', _VER_C60)
  include('cwsynchc.inc'),once
  ****
  include('fm3equ.clw')
ds_FM_Upgrading  &byte,thread                ! File Manager 2/3 upgrading flag
Access:Projects      &FileManager,THREAD                   ! FileManager for Projects
Relate:Projects      &RelationManager,THREAD               ! RelationManager for Projects

FuzzyMatcher         FuzzyClass                            ! Global fuzzy matcher
GlobalErrorStatus    ErrorStatusClass,THREAD
GlobalErrors         ErrorClass                            ! Global error manager
INIMgr               INIClass                              ! Global non-volatile storage manager
GlobalRequest        BYTE(0),THREAD                        ! Set when a browse calls a form, to let it know action to perform
GlobalResponse       BYTE(0),THREAD                        ! Set to the response from the form
VCRRequest           LONG(0),THREAD                        ! Set to the request from the VCR buttons

Dictionary           CLASS,THREAD
Construct              PROCEDURE
Destruct               PROCEDURE
                     END


  CODE
  GlobalErrors.Init(GlobalErrorStatus)
  FuzzyMatcher.Init                                        ! Initilaize the browse 'fuzzy matcher'
  FuzzyMatcher.SetOption(MatchOption:NoCase, 1)            ! Configure case matching
  FuzzyMatcher.SetOption(MatchOption:WordOnly, 0)          ! Configure 'word only' matching
  INIMgr.Init('.\HtmlPageScraper.INI', NVD_INI)            ! Configure INIManager to use INI file
  DctInit()
                             ! Begin Generated by NetTalk Extension Template
  
    if ~command ('/netnolog') and (command ('/nettalklog') or command ('/nettalklogerrors') or command ('/neterrors') or command ('/netall'))
      NetDebugTrace ('[Nettalk Template] NetTalk Template version 14.23')
      NetDebugTrace ('[Nettalk Template] NetTalk Template using Clarion ' & 11000)
      NetDebugTrace ('[Nettalk Template] NetTalk Object version ' & NETTALK:VERSION )
      NetDebugTrace ('[Nettalk Template] ABC Template Chain')
    end
                             ! End Generated by Extension Template
  SYSTEM{PROP:Icon} = 'AppIcon.ico'
                 !CapeSoft MessageBox init code
  ThisMessageBox.init(1,1)
                 !End of CapeSoft MessageBox init code
   ! Generated using Clarion Template version v11.0  Family = abc
    ds_FM_Upgrading &= ds_PassHandleForUpgrading()
    if ds_FMInited = 0
      ds_FMInited = 1
        ds_SetOption('inifilename','.\fm3.ini')
      ds_SetOption('BadFile',0)
      ds_AddDriver('Tps',gTopSpeedFile,__gtps:record)
  
  
        ! Thisapp = 0
      ds_IgnoreDriver('AllFiles',1)
      ds_SetOption('SPCreate',1)
      ds_SetOption('GUIDsCaseInsensitive',1)
    ds_UsingFileEx('Projects',Projects,2+ds_VersionModifier,'Pro')
              omit('***',FM2=1)
              !! Don't forget to add the FM2=>1 define to your project
              You did forget didn't you ?
              ! close this window - go to the app - click on project - click on properties -
              ! click on the defines tab - add FM2=>1 to the defines...
            !***
    End  !End of if ds_FMInited = 0
    ds_SetOKToEndSessionHandler(address(MyOKToEndSessionHandler))
    ds_SetEndSessionHandler(address(MyEndSessionHandler))
  Main
  INIMgr.Update
    ThisGPF.RestartProgram = 0
                             ! Begin Generated by NetTalk Extension Template
    NetCloseCallBackWindow() ! Tell NetTalk DLL to shutdown it's WinSock Call Back Window
  
    if ~command ('/netnolog') and (command ('/nettalklog') or command ('/nettalklogerrors') or command ('/neterrors') or command ('/netall'))
      NetDebugTrace ('[Nettalk Template] NetTalk Template version 14.23')
      NetDebugTrace ('[Nettalk Template] NetTalk Template using Clarion ' & 11000)
      NetDebugTrace ('[Nettalk Template] Closing Down NetTalk (Object) version ' & NETTALK:VERSION)
    end
                             ! End Generated by Extension Template
      ThisMessageBox.Kill()                     !CapeSoft MessageBox template generated code
  INIMgr.Kill                                              ! Destroy INI manager
  FuzzyMatcher.Kill                                        ! Destroy fuzzy matcher
    
!----------------------------------------------------
ThisGPF.Construct     PROCEDURE ()
  CODE
!----------------------------------------------------
ThisGPF.Destruct     PROCEDURE ()
  CODE
!----------------------------------------------------
ThisGPF._GetSymbol     PROCEDURE (ulong pAddress,byte pStackTrace=1)
ReturnValue   any
  CODE
  ReturnValue = PARENT._GetSymbol (pAddress,pStackTrace)
    Return ReturnValue
!----------------------------------------------------
ThisGPF._LookupExceptionCode     PROCEDURE (ulong p_ExceptionCode)
ReturnValue   any
  CODE
  ReturnValue = PARENT._LookupExceptionCode (p_ExceptionCode)
    Return ReturnValue
!----------------------------------------------------
ThisGPF._VectoredExceptionHandler_     PROCEDURE (ulong p_e)
ReturnValue   long
  CODE
  ReturnValue = PARENT._VectoredExceptionHandler_ (p_e)
    Return ReturnValue
!----------------------------------------------------
ThisGPF._StackDetails     PROCEDURE (ulong p_e,byte p_details,ulong p_hProcess)
ReturnValue   any
  CODE
  ReturnValue = PARENT._StackDetails (p_e,p_details,p_hProcess)
    Return ReturnValue
!----------------------------------------------------
ThisGPF._LocateDebugSymbols     PROCEDURE (long phModule)
ReturnValue   byte
  CODE
  ReturnValue = PARENT._LocateDebugSymbols (phModule)
    Return ReturnValue
!----------------------------------------------------
ThisGPF._ReadBlockFromFile     PROCEDURE (ulong pOffset,long pReadBytes,*string pFileBlock,*string pFileName)
  CODE
  PARENT._ReadBlockFromFile (pOffset,pReadBytes,pFileBlock,pFileName)
!----------------------------------------------------
ThisGPF._GetModuleName     PROCEDURE (long phModule)
ReturnValue   any
  CODE
  ReturnValue = PARENT._GetModuleName (phModule)
    Return ReturnValue
!----------------------------------------------------
ThisGPF._GetModuleHandle     PROCEDURE (ulong pAddress)
ReturnValue   long
  CODE
  ReturnValue = PARENT._GetModuleHandle (pAddress)
    Return ReturnValue
!----------------------------------------------------
ThisGPF.LookupAddress     PROCEDURE ()
  CODE
  PARENT.LookupAddress ()
!----------------------------------------------------
ThisGPF.Initialize     PROCEDURE ()
  CODE
  ThisGPF.EmailAddress = 'The developer <support@example.com>'
  ThisGPF.WindowTitle = ''
  ThisGPF.VersionNumber = ds_GetFileVersionInfo()
  ThisGPF.DumpFileName = 'GPFReport.log'
  ThisGPF.AllowEmail = 1
  ThisGPF.DumpFileAppend = 1
  ThisGPF.RestartProgram = 0
  ThisGPF.ShowDetails = 0
  ThisGPF.DebugEmail = 0
  ThisGPF.DebugLogEnabled = 0
  ThisGPF.WaitWinEnabled = 0
  ThisGPF.Workstation = ds_GetWorkstationName()     ! requires winevent ver 3.61 or later
  ThisGPF.UserName = ds_GetUserName()               ! requires winevent ver 3.61 or later
  PARENT.Initialize ()
!----------------------------------------------------
ThisGPF.FilterExceptions     PROCEDURE (ulong pException)
  CODE
  PARENT.FilterExceptions (pException)
!----------------------------------------------------
ThisGPF._EncodeEmail     PROCEDURE (string pEmailText)
ReturnValue   any
  CODE
  ReturnValue = PARENT._EncodeEmail (pEmailText)
    Return ReturnValue
!----------------------------------------------------
ThisGPF._FindFirstBreak     PROCEDURE (string pText,long pMaxLen)
ReturnValue   long
  CODE
  ReturnValue = PARENT._FindFirstBreak (pText,pMaxLen)
    Return ReturnValue
!----------------------------------------------------
ThisGPF.ExtraReportText     PROCEDURE ()
  CODE
  ! ThisGPF.ReportText = 'Add your own report text here.'<13,10>This is on the next line.'
  PARENT.ExtraReportText ()
!----------------------------------------------------
ThisGPF._GetDLLVersion     PROCEDURE (string pDLLName)
ReturnValue   any
  CODE
  ReturnValue = PARENT._GetDLLVersion (pDLLName)
    Return ReturnValue
!----------------------------------------------------
ThisGPF._SetFileNames     PROCEDURE ()
  CODE
  PARENT._SetFileNames ()
!----------------------------------------------------
ThisGPF.DeleteDumpFile     PROCEDURE ()
  CODE
  PARENT.DeleteDumpFile ()
!----------------------------------------------------
ThisGPF._InitReportText     PROCEDURE ()
  CODE
  PARENT._InitReportText ()
!----------------------------------------------------
ThisGPF._ExecuteCommands     PROCEDURE ()
  CODE
  PARENT._ExecuteCommands ()
!----------------------------------------------------
ThisGPF._StackDump     PROCEDURE (long pStart,long pEnd,long pStackLevel)
ReturnValue   any
  CODE
  ReturnValue = PARENT._StackDump (pStart,pEnd,pStackLevel)
    Return ReturnValue
!----------------------------------------------------
ThisGPF._FindLinePosition     PROCEDURE (string pText,long pLineNumber)
ReturnValue   long
  CODE
  ReturnValue = PARENT._FindLinePosition (pText,pLineNumber)
    Return ReturnValue
!----------------------------------------------------
ThisGPF._DebugLog     PROCEDURE (string pDebugData,byte pFirstLine=0)
  CODE
  PARENT._DebugLog (pDebugData,pFirstLine)
!----------------------------------------------------
ThisGPF._FormatLineInfo     PROCEDURE (long pLineNumber,string pProcName,string pSourceName,string pModuleName,byte pStackTrace,byte pNoProcFound,byte pExactAddress,byte pNoLineNumber)
ReturnValue   any
  CODE
  ReturnValue = PARENT._FormatLineInfo (pLineNumber,pProcName,pSourceName,pModuleName,pStackTrace,pNoProcFound,pExactAddress,pNoLineNumber)
    Return ReturnValue
!----------------------------------------------------
ThisGPF._GetAssert     PROCEDURE (long pSP,long pBP)
ReturnValue   long
  CODE
  ReturnValue = PARENT._GetAssert (pSP,pBP)
    Return ReturnValue
!----------------------------------------------------
ThisGPF._GetOtherMessage     PROCEDURE (long pSP,long pBP)
ReturnValue   long
  CODE
  ReturnValue = PARENT._GetOtherMessage (pSP,pBP)
    Return ReturnValue
!----------------------------------------------------
ThisGPF._RestartProgram     PROCEDURE ()
  CODE
  PARENT._RestartProgram ()
ThisMessageBox.AssignGlobalClass        procedure     !CapeSoft MessageBox Object Procedure

  Code
    parent.AssignGlobalClass 
    self.GlobalClass &= ThisMessageBoxGlobal

ThisMessageBox.Init                     procedure  (long UseABCClasses=0,long UseDefaultFile=0)   !CapeSoft MessageBox Object Procedure
TempVar         long,dim(2)
TMPLogFileName  string(252)

  Code
    parent.Init (UseABCClasses,UseDefaultFile)
    system{prop:MessageHook} = address(ds_Message)
    system{prop:StopHook} = address(ds_Stop)
    system{prop:HaltHook} = address(ds_Halt)
  self.SetGlobalSetting('LogMessages',1)
  self.SetGlobalSetting('ShowTimeOut',0)
  self.SetGlobalSetting('GPFHotKey', CtrlAltG)
  self.SetGlobalSetting('CopyKey', CtrlC)
    self.SetGlobalSetting('INIFile','CSMesBox.INI')
    self.SetGlobalSetting('INISection', 'CS_Messages')
    TMPLogFileName = command('0')
    loop
      EndSearchEXEName = instring('\',TMPLogFileName,1,StartSearchEXEName)
      if ~EndSearchEXEName then break .
      StartSearchEXEName = EndSearchEXEName + 1
    end
    if StartSearchEXEName > 1
      TMPLogFileName = TMPLogFileName[1:(StartSearchEXEName - 1)]
    else
      TMPLogFileName = path() & '\'
    end
    self.SetGlobalSetting('LogFileName',clip(TMPLogFileName) & '\MessageBox.Log')
    self.SetGlobalSetting('DateFormat', '@d17')

ThisMessageBox.PrimeLog                 procedure  (<string ExtraDetails>)   !CapeSoft MessageBox Object Procedure

  Code
              !CapeSoft MessageBox Template Generated code to prime the logging records
                 !Because there is no File selected, record priming is handled in the object and
                 !the default ASCII file is used to log the records.
    parent.PrimeLog (ExtraDetails)

! ------ winevent -------------------------------------------------------------------
MyOKToEndSessionHandler procedure(long pLogoff)
OKToEndSession    long(TRUE)
! Setting the return value OKToEndSession = FALSE
! will tell windows not to shutdown / logoff now.
! If parameter pLogoff = TRUE if the user is logging off.

  code
  return(OKToEndSession)

! ------ winevent -------------------------------------------------------------------
MyEndSessionHandler procedure(long pLogoff)
! If parameter pLogoff = TRUE if the user is logging off.

  code


Dictionary.Construct PROCEDURE

  CODE
  IF THREAD()<>1
     DctInit()
  END


Dictionary.Destruct PROCEDURE

  CODE
  DctKill()

