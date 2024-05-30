

   MEMBER('HtmlPageScraper.clw')                           ! This is a MEMBER module


   INCLUDE('ABTOOLBA.INC'),ONCE
   INCLUDE('ABUTIL.INC'),ONCE
   INCLUDE('ABWINDOW.INC'),ONCE

                     MAP
                       INCLUDE('HTMLPAGESCRAPER002.INC'),ONCE        !Local module procedure declarations
                     END


!!! <summary>
!!! Generated from procedure template - Window
!!! Form Projects
!!! </summary>
UpdateProjects PROCEDURE 

CurrentTab           STRING(80)                            ! 
ActionMessage        CSTRING(40)                           ! 
History::Pro:Record  LIKE(Pro:RECORD),THREAD
QuickWindow          WINDOW('Form Projects'),AT(,,351,80),FONT('Segoe UI',10,COLOR:Black,FONT:regular,CHARSET:DEFAULT), |
  RESIZE,AUTO,CENTER,ICON('AppIcon.ico'),GRAY,IMM,HLP('UpdateProjects'),SYSTEM,WALLPAPER('Video Capt' & |
  'ure_Gradient_20240530_121419.png')
                       BUTTON('&OK'),AT(241,64,49,14),USE(?OK),LEFT,ICON('Check.ico'),DEFAULT,FLAT,MSG('Accept dat' & |
  'a and close the window'),TIP('Accept data and close the window')
                       BUTTON('&Cancel'),AT(294,64,49,14),USE(?Cancel),LEFT,ICON('Close.ico'),FLAT,MSG('Cancel operation'), |
  TIP('Cancel operation')
                       ENTRY(@s50),AT(57,4,204,10),USE(Pro:Description),REQ
                       PROMPT('Description:'),AT(6,4),USE(?Pro:Description:Prompt),TRN
                       PROMPT('Main URL:'),AT(6,18),USE(?Pro:MainUrl:Prompt),TRN
                       ENTRY(@s100),AT(57,18,271,10),USE(Pro:MainUrl),REQ
                       PROMPT('Save To:'),AT(6,32,28,10),USE(?Pro:SaveTo:Prompt),TRN
                       TEXT,AT(57,32,271,23),USE(Pro:SaveTo),REQ
                       BUTTON,AT(331,32,15,12),USE(?LookupFile),ICON('Search.ico'),FLAT
                       BUTTON,AT(331,16,15,12),USE(?BrowserBtn),ICON('Earth.ico'),FLAT,HIDE
                     END

    omit('***',WE::CantCloseNowSetHereDone=1)  !Getting Nested omit compile error, then uncheck the "Check for duplicate CantCloseNowSetHere variable declaration" in the WinEvent local template
WE::CantCloseNowSetHereDone equate(1)
WE::CantCloseNowSetHere     long
    !***
ThisWindow           CLASS(WindowManager)
Ask                    PROCEDURE(),DERIVED
Init                   PROCEDURE(),BYTE,PROC,DERIVED
Kill                   PROCEDURE(),BYTE,PROC,DERIVED
Run                    PROCEDURE(),BYTE,PROC,DERIVED
TakeAccepted           PROCEDURE(),BYTE,PROC,DERIVED
TakeEvent              PROCEDURE(),BYTE,PROC,DERIVED
TakeWindowEvent        PROCEDURE(),BYTE,PROC,DERIVED
                     END

Toolbar              ToolbarClass
! ----- ThisHyperActive --------------------------------------------------------------------------
ThisHyperActive      Class(csHyperActiveClass)
    ! derived method declarations
GetListURL             PROCEDURE (LONG HyperControl,STRING HyperURL,<STRING FirstEmailName>,<STRING LastEmailName>),STRING ,VIRTUAL
RefreshURL             PROCEDURE (LONG HyperControl),long ,VIRTUAL
                     End  ! ThisHyperActive
! ----- end ThisHyperActive -----------------------------------------------------------------------
! ----- csResize --------------------------------------------------------------------------
csResize             Class(csResizeClass)
    ! derived method declarations
Fetch                  PROCEDURE (STRING Sect,STRING Ent,*? Val),VIRTUAL
Update                 PROCEDURE (STRING Sect,STRING Ent,STRING Val),VIRTUAL
Init                   PROCEDURE (),VIRTUAL
                     End  ! csResize
! ----- end csResize -----------------------------------------------------------------------
FileLookup8          SelectFileClass
CurCtrlFeq          LONG
FieldColorQueue     QUEUE
Feq                   LONG
OldColor              LONG
                    END

  CODE
? DEBUGHOOK(Projects:Record)
  GlobalResponse = ThisWindow.Run()                        ! Opens the window and starts an Accept Loop

!---------------------------------------------------------------------------
DefineListboxStyle ROUTINE
!|
!| This routine create all the styles to be shared in this window
!| It`s called after the window open
!|
!---------------------------------------------------------------------------

ThisWindow.Ask PROCEDURE

  CODE
  CASE SELF.Request                                        ! Configure the action message text
  OF ViewRecord
    ActionMessage = 'View Record'
  OF InsertRecord
    ActionMessage = 'Record Will Be Added'
  OF ChangeRecord
    ActionMessage = 'Record Will Be Changed'
  END
  QuickWindow{PROP:Text} = ActionMessage                   ! Display status message in title bar
  CASE SELF.Request
  OF ChangeRecord OROF DeleteRecord
    QuickWindow{PROP:Text} = QuickWindow{PROP:Text} & '  (' & Clip(Pro:Description) & ')' ! Append status message to window title text
  OF InsertRecord
    QuickWindow{PROP:Text} = QuickWindow{PROP:Text} & '  (New)'
  END
  PARENT.Ask


ThisWindow.Init PROCEDURE

ReturnValue          BYTE,AUTO

  CODE
  GlobalErrors.SetProcedureName('UpdateProjects')
  SELF.Request = GlobalRequest                             ! Store the incoming request
  ReturnValue = PARENT.Init()
  IF ReturnValue THEN RETURN ReturnValue.
  SELF.FirstField = ?OK
  SELF.VCRRequest &= VCRRequest
  SELF.Errors &= GlobalErrors                              ! Set this windows ErrorManager to the global ErrorManager
  CLEAR(GlobalRequest)                                     ! Clear GlobalRequest after storing locally
  CLEAR(GlobalResponse)
  SELF.AddItem(Toolbar)
  SELF.HistoryKey = CtrlH
  SELF.AddHistoryFile(Pro:Record,History::Pro:Record)
  SELF.AddHistoryField(?Pro:Description,2)
  SELF.AddHistoryField(?Pro:MainUrl,3)
  SELF.AddHistoryField(?Pro:SaveTo,4)
  SELF.AddUpdateFile(Access:Projects)
  SELF.AddItem(?Cancel,RequestCancelled)                   ! Add the cancel control to the window manager
  Relate:Projects.Open()                                   ! File Projects used by this procedure, so make sure it's RelationManager is open
  SELF.FilesOpened = True
  SELF.Primary &= Relate:Projects
  IF SELF.Request = ViewRecord AND NOT SELF.BatchProcessing ! Setup actions for ViewOnly Mode
    SELF.InsertAction = Insert:None
    SELF.DeleteAction = Delete:None
    SELF.ChangeAction = Change:None
    SELF.CancelAction = Cancel:Cancel
    SELF.OkControl = 0
  ELSE
    SELF.ChangeAction = Change:Caller                      ! Changes allowed
    SELF.CancelAction = Cancel:Cancel+Cancel:Query         ! Confirm cancel
    SELF.OkControl = ?OK
    IF SELF.PrimeUpdate() THEN RETURN Level:Notify.
  END
  SELF.Open(QuickWindow)                                   ! Open window
  Do DefineListboxStyle
  Alert(AltKeyPressed)  ! WinEvent : These keys cause a program to crash on Windows 7 and Windows 10.
  Alert(F10Key)         !
  Alert(CtrlF10)        !
  Alert(ShiftF10)       !
  Alert(CtrlShiftF10)   !
  Alert(AltSpace)       !
  WinAlertMouseZoom()
  WinAlert(WE::WM_QueryEndSession,,Return1+PostUser)
  IF SELF.Request = ViewRecord                             ! Configure controls for View Only mode
    ?Pro:Description{PROP:ReadOnly} = True
    ?Pro:MainUrl{PROP:ReadOnly} = True
    DISABLE(?LookupFile)
    DISABLE(?BrowserBtn)
  END
  csResize.Init('UpdateProjects',QuickWindow,1)
  INIMgr.Fetch('UpdateProjects',QuickWindow)               ! Restore window settings from non-volatile store
  FileLookup8.Init
  FileLookup8.ClearOnCancel = True
  FileLookup8.Flags=BOR(FileLookup8.Flags,FILE:LongName)   ! Allow long filenames
  FileLookup8.Flags=BOR(FileLookup8.Flags,FILE:Directory)  ! Allow Directory Dialog
  FileLookup8.SetMask('All Files','*.*')                   ! Set the file mask
  FileLookup8.DefaultDirectory='Longpath() & ''\Extracted Files'''
                   ! Start of HyperActive Init Code
  ThisHyperActive.HandCursorToUse = 'hand2.cur'
  ThisHyperActive.HandCursorToUse = '~' & ThisHyperActive.HandCursorToUse
  0{prop:timer} = 8
  ThisHyperActive.Init(event:Timer )
  ThisHyperActive.LimitURL=-1
  ThisHyperActive.ItemQ.SkypeFunction = 0
  ThisHyperActive.AddItem(?BrowserBtn,Pro:MainUrl)
  ThisHyperActive.Refresh()
                   ! End of HyperActive Init Code
  csResize.Open()
  SELF.SetAlerts()
    If Pro:MainUrl
        UNHIDE(?BrowserBtn)
    ELSE
        HIDE(?BrowserBtn)
    End  
  RETURN ReturnValue


ThisWindow.Kill PROCEDURE

ReturnValue          BYTE,AUTO

  CODE
  If self.opened Then WinAlert().
  ReturnValue = PARENT.Kill()
  IF ReturnValue THEN RETURN ReturnValue.
  IF SELF.FilesOpened
    Relate:Projects.Close()
  END
  IF SELF.Opened
    INIMgr.Update('UpdateProjects',QuickWindow)            ! Save window data to non-volatile store
  END
  GlobalErrors.SetProcedureName
  RETURN ReturnValue


ThisWindow.Run PROCEDURE

ReturnValue          BYTE,AUTO

  CODE
  ReturnValue = PARENT.Run()
  IF SELF.Request = ViewRecord                             ! In View Only mode always signal RequestCancelled
    ReturnValue = RequestCancelled
  END
  RETURN ReturnValue


ThisWindow.TakeAccepted PROCEDURE

ReturnValue          BYTE,AUTO

Looped BYTE
  CODE
  LOOP                                                     ! This method receive all EVENT:Accepted's
    IF Looped
      RETURN Level:Notify
    ELSE
      Looped = 1
    END
  ReturnValue = PARENT.TakeAccepted()
    CASE ACCEPTED()
    OF ?OK
      ThisWindow.Update()
      IF SELF.Request = ViewRecord AND NOT SELF.BatchProcessing THEN
         POST(EVENT:CloseWindow)
      END
    OF ?Pro:Description
      Pro:SaveTo = Longpath() & '\Extracted Files\' & Clip(Pro:Description) 
    OF ?LookupFile
      ThisWindow.Update()
      Pro:SaveTo = FileLookup8.Ask(1)
      DISPLAY
      Pro:SaveTo = Clip(Pro:SaveTo) & '\' & Clip(Pro:Description)
    END
    RETURN ReturnValue
  END
  ReturnValue = Level:Fatal
  RETURN ReturnValue


ThisWindow.TakeEvent PROCEDURE

ReturnValue          BYTE,AUTO

Looped BYTE
  CODE
  csResize.TakeEvent()
  LOOP                                                     ! This method receives all events
    IF Looped
      RETURN Level:Notify
    ELSE
      Looped = 1
    END
    Case EVENT()
    Of EVENT:Accepted
        If Pro:MainUrl
            UNHIDE(?BrowserBtn)
        ELSE
            HIDE(?BrowserBtn)
        End
    End  
    ThisHyperActive.TakeEvent()                      !HyperActive Code
  ReturnValue = PARENT.TakeEvent()
  If event() = event:VisibleOnDesktop !or event() = event:moved
    ds_VisibleOnDesktop()
  end
    RETURN ReturnValue
  END
  ReturnValue = Level:Fatal
  RETURN ReturnValue


ThisWindow.TakeWindowEvent PROCEDURE

ReturnValue          BYTE,AUTO

Looped BYTE
  CODE
  LOOP                                                     ! This method receives all window specific events
    IF Looped
      RETURN Level:Notify
    ELSE
      Looped = 1
    END
    CASE EVENT()
    OF EVENT:CloseDown
      if WE::CantCloseNow
        WE::MustClose = 1
        cycle
      else
        self.CancelAction = cancel:cancel
        self.response = requestcancelled
      end
    END
  ReturnValue = PARENT.TakeWindowEvent()
    CASE EVENT()
    OF EVENT:OpenWindow
      ThisHyperActive.Refresh()
        post(event:visibleondesktop)
    END
    RETURN ReturnValue
  END
  ReturnValue = Level:Fatal
  RETURN ReturnValue

!----------------------------------------------------
ThisHyperActive.GetListURL   PROCEDURE (LONG HyperControl,STRING HyperURL,<STRING FirstEmailName>,<STRING LastEmailName>)
ReturnValue   any
HATempVar                 long
  CODE
  Return ReturnValue
  ReturnValue = PARENT.GetListURL (HyperControl,HyperURL,FirstEmailName,LastEmailName)
    Return ReturnValue
!----------------------------------------------------
ThisHyperActive.RefreshURL   PROCEDURE (LONG HyperControl)
ReturnValue   long
  CODE
      if ~omitted(2)
        case HyperControl
        of ?BrowserBtn
          self.URLTORun = Pro:MainUrl
          return(1)
        end
      end
  ReturnValue = PARENT.RefreshURL (HyperControl)
    Return ReturnValue
!----------------------------------------------------
csResize.Fetch   PROCEDURE (STRING Sect,STRING Ent,*? Val)
  CODE
  INIMgr.Fetch(Sect,Ent,Val)
  PARENT.Fetch (Sect,Ent,Val)
!----------------------------------------------------
csResize.Update   PROCEDURE (STRING Sect,STRING Ent,STRING Val)
  CODE
  INIMgr.Update(Sect,Ent,Val)
  PARENT.Update (Sect,Ent,Val)
!----------------------------------------------------
csResize.Init   PROCEDURE ()
  CODE
  PARENT.Init ()
  Self.CornerStyle = Ras:CornerDots
  SELF.GrabCornerLines() !
  SELF.SetStrategy(?OK,100,100,0,0)
  SELF.SetStrategy(?Cancel,100,100,0,0)
  SELF.SetStrategy(?Pro:SaveTo,0,0,100,100)
  SELF.SetStrategy(?LookupFile,100,,0,)
