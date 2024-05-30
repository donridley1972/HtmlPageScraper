

   MEMBER('HtmlPageScraper.clw')                           ! This is a MEMBER module


   INCLUDE('ABBROWSE.INC'),ONCE
   INCLUDE('ABPOPUP.INC'),ONCE
   INCLUDE('ABTOOLBA.INC'),ONCE
   INCLUDE('ABWINDOW.INC'),ONCE
   INCLUDE('NetWww.inc'),ONCE

                     MAP
                       INCLUDE('HTMLPAGESCRAPER001.INC'),ONCE        !Local module procedure declarations
                       INCLUDE('HTMLPAGESCRAPER002.INC'),ONCE        !Req'd for module callout resolution
                     END


!!! <summary>
!!! Generated from procedure template - Window
!!! </summary>
Main PROCEDURE 

MyState             Long
MyState:Start       Equate(1)
MyState:GetPages    Equate(2)
MyState:GetPage     Equate(3)
MyState:WaitOnPage  Equate(4)
MyState:ParsePage   Equate(5)
MyState:GetFile     Equate(6)
MyState:WaitOnFile  Equate(7)
FileNdx              LONG                                  ! 
PagesQ               QUEUE,PRE(Pag)                        ! 
PageLink             STRING(256)                           ! 
                     END                                   ! 
LinksQ               QUEUE,PRE(Lin)                        ! 
Link                 STRING(1000)                          ! 
FileName             STRING(256)                           ! 
FileExt              STRING(5)                             ! 
                     END                                   ! 
BRW4::View:Browse    VIEW(Projects)
                       PROJECT(Pro:Description)
                       PROJECT(Pro:MainUrl)
                       PROJECT(Pro:SaveTo)
                       PROJECT(Pro:Guid)
                     END
Queue:Browse         QUEUE                            !Queue declaration for browse/combo box using ?ListProjects
Pro:Description        LIKE(Pro:Description)          !List box control field - type derived from field
Pro:MainUrl            LIKE(Pro:MainUrl)              !Browse hot field - type derived from field
Pro:SaveTo             LIKE(Pro:SaveTo)               !Browse hot field - type derived from field
Pro:Guid               LIKE(Pro:Guid)                 !Primary key field - type derived from field
Mark                   BYTE                           !Entry's marked status
ViewPosition           STRING(1024)                   !Entry's view position
                     END
Window               WINDOW('HTML Page Scraper'),AT(,,587,270),FONT('Segoe UI',10),RESIZE,AUTO,ICON('AppIcon.ico'), |
  GRAY,SYSTEM,WALLPAPER('Video Capture_Gradient_20240530_121419.png'),IMM
                       BUTTON('Close'),AT(539,248,46,14),USE(?Close),LEFT,ICON('Exit.ico'),FLAT
                       LIST,AT(3,53,98,190),USE(?ListProjects),LEFT(2),VSCROLL,FORMAT('200L(2)|M~Projects~@s50@'), |
  FROM(Queue:Browse),IMM,TIP('Projects List')
                       BUTTON,AT(105,53,22,12),USE(?Insert),ICON('Add.ico'),FLAT,TIP('New Project')
                       BUTTON,AT(105,67,22,12),USE(?Change),ICON('Edit.ico'),FLAT,TIP('Edit Project')
                       BUTTON,AT(105,82,22,12),USE(?Delete),ICON('Garbage.ico'),FLAT,TIP('Delete Project')
                       BUTTON,AT(105,96,22,12),USE(?StartBtn),ICON('ArrowRight1.ico'),FLAT,TIP('Start Scraping')
                       SHEET,AT(130,53,454,192),USE(?SHEET1)
                         TAB('Page Links'),USE(?TAB2),ICON('Link.ico')
                           LIST,AT(135,70,445,171),USE(?ListPageLinks),HVSCROLL,FORMAT('1020L(2)M~Page Link~@s255@'), |
  FROM(PagesQ),TIP('Page Links List')
                         END
                         TAB('File Links'),USE(?TAB1),ICON('Files.ico')
                           LIST,AT(135,70,445,171),USE(?ListFileLinks),HVSCROLL,FORMAT('212L(2)M~Link~@s255@125L(2' & |
  ')M~File Name~@s255@20L(2)M~File Ext~@s5@'),FROM(LinksQ),TIP('File Links List')
                         END
                       END
                       BUTTON,AT(105,111,22,12),USE(?StopBtn),ICON('MediaStop1.ico'),FLAT,TIP('Stop Scraping')
                       IMAGE('Logo.png'),AT(2,1,295,48),USE(?IMAGE1)
                     END

    omit('***',WE::CantCloseNowSetHereDone=1)  !Getting Nested omit compile error, then uncheck the "Check for duplicate CantCloseNowSetHere variable declaration" in the WinEvent local template
WE::CantCloseNowSetHereDone equate(1)
WE::CantCloseNowSetHere     long
    !***
InSt                StringTheory
LneSt               StringTheory
SrcSt               StringTheory
SrcSetSt            StringTheory
MenuSt              StringTheory
FullPageSt          StringTheory
TempSt              StringTheory
FileSt              StringTheory
x                   Long
y                   Long
MyTrace             dwrTrace

local                   CLASS
RemoveSection           Procedure(StringTheory pSt,string pTag,long pExclusive) !,string
                        End
ThisWindow           CLASS(WindowManager)
Init                   PROCEDURE(),BYTE,PROC,DERIVED
Kill                   PROCEDURE(),BYTE,PROC,DERIVED
Run                    PROCEDURE(USHORT Number,BYTE Request),BYTE,PROC,DERIVED
TakeAccepted           PROCEDURE(),BYTE,PROC,DERIVED
TakeEvent              PROCEDURE(),BYTE,PROC,DERIVED
TakeWindowEvent        PROCEDURE(),BYTE,PROC,DERIVED
                     END

Toolbar              ToolbarClass
! ----- csResize --------------------------------------------------------------------------
csResize             Class(csResizeClass)
    ! derived method declarations
Fetch                  PROCEDURE (STRING Sect,STRING Ent,*? Val),VIRTUAL
Update                 PROCEDURE (STRING Sect,STRING Ent,STRING Val),VIRTUAL
Init                   PROCEDURE (),VIRTUAL
                     End  ! csResize
! ----- end csResize -----------------------------------------------------------------------
BrwProjects          CLASS(BrowseClass)                    ! Browse using ?ListProjects
Q                      &Queue:Browse                  !Reference to browse queue
Init                   PROCEDURE(SIGNED ListBox,*STRING Posit,VIEW V,QUEUE Q,RelationManager RM,WindowManager WM)
                     END

!Local Data Classes
ThisWebClient        CLASS(NetWebClient)                   ! Generated by NetTalk Extension (Class Definition)
ErrorTrap              PROCEDURE(string errorStr,string functionName),DERIVED
PageReceived           PROCEDURE(),DERIVED

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

ThisWindow.Init PROCEDURE

ReturnValue          BYTE,AUTO

  CODE
  GlobalErrors.SetProcedureName('Main')
  SELF.Request = GlobalRequest                             ! Store the incoming request
  ReturnValue = PARENT.Init()
  IF ReturnValue THEN RETURN ReturnValue.
  SELF.FirstField = ?Close
  SELF.VCRRequest &= VCRRequest
  SELF.Errors &= GlobalErrors                              ! Set this windows ErrorManager to the global ErrorManager
  CLEAR(GlobalRequest)                                     ! Clear GlobalRequest after storing locally
  CLEAR(GlobalResponse)
  SELF.AddItem(Toolbar)
  IF SELF.Request = SelectRecord
     SELF.AddItem(?Close,RequestCancelled)                 ! Add the close control to the window manger
  ELSE
     SELF.AddItem(?Close,RequestCompleted)                 ! Add the close control to the window manger
  END
  Relate:Projects.Open()                                   ! File Projects used by this procedure, so make sure it's RelationManager is open
  SELF.FilesOpened = True
  BrwProjects.Init(?ListProjects,Queue:Browse.ViewPosition,BRW4::View:Browse,Queue:Browse,Relate:Projects,SELF) ! Initialize the browse manager
  SELF.Open(Window)                                        ! Open window
    Window{PROP:Pixels} = 1
                                               ! Generated by NetTalk Extension (Start)
  ThisWebClient.SuppressErrorMsg = 1         ! No Object Generated Error Messages ! Generated by NetTalk Extension
  ThisWebClient.init()
  if ThisWebClient.error <> 0
    ! Put code in here to handle if the object does not initialise properly
  end
  Do DefineListboxStyle
  Alert(AltKeyPressed)  ! WinEvent : These keys cause a program to crash on Windows 7 and Windows 10.
  Alert(F10Key)         !
  Alert(CtrlF10)        !
  Alert(ShiftF10)       !
  Alert(CtrlShiftF10)   !
  Alert(AltSpace)       !
  WinAlertMouseZoom()
  WinAlert(WE::WM_QueryEndSession,,Return1+PostUser)
  BrwProjects.Q &= Queue:Browse
  BrwProjects.AddSortOrder(,)                              ! Add the sort order for  for sort order 1
  BrwProjects.AddField(Pro:Description,BrwProjects.Q.Pro:Description) ! Field Pro:Description is a hot field or requires assignment from browse
  BrwProjects.AddField(Pro:MainUrl,BrwProjects.Q.Pro:MainUrl) ! Field Pro:MainUrl is a hot field or requires assignment from browse
  BrwProjects.AddField(Pro:SaveTo,BrwProjects.Q.Pro:SaveTo) ! Field Pro:SaveTo is a hot field or requires assignment from browse
  BrwProjects.AddField(Pro:Guid,BrwProjects.Q.Pro:Guid)    ! Field Pro:Guid is a hot field or requires assignment from browse
  csResize.Init('Main',Window,1)
  INIMgr.Fetch('Main',Window)                              ! Restore window settings from non-volatile store
  BrwProjects.AskProcedure = 1                             ! Will call: UpdateProjects
  BrwProjects.AddToolbarTarget(Toolbar)                    ! Browse accepts toolbar control
  csResize.Open()
  SELF.SetAlerts()
  !ThisWebClient.NetShowSend     =  1
  !ThisWebClient.NetShowReceive  =  1
    If Not EXISTS(LONGPATH() & '\Extracted Files')
        ds_CreateDirectory(LONGPATH() & '\Extracted Files')
    End
    ?ListProjects{PROP:LineHeight} = 25
    ?ListPageLinks{PROP:LineHeight} = 25
    ?ListFileLinks{PROP:LineHeight} = 25
    MyTrace.Trace('')
    MyTrace.Trace('Image Width = ' & ?IMAGE1{PROP:Width})
    MyTrace.Trace('Image Height = ' & ?IMAGE1{PROP:Height})
    MyTrace.Trace('')
    MyTrace.Trace('')
  RETURN ReturnValue


ThisWindow.Kill PROCEDURE

ReturnValue          BYTE,AUTO

  CODE
  ThisWebClient.Kill()                              ! Generated by NetTalk Extension
  If self.opened Then WinAlert().
  ReturnValue = PARENT.Kill()
  IF ReturnValue THEN RETURN ReturnValue.
  IF SELF.FilesOpened
    Relate:Projects.Close()
  END
  IF SELF.Opened
    INIMgr.Update('Main',Window)                           ! Save window data to non-volatile store
  END
  GlobalErrors.SetProcedureName
  RETURN ReturnValue


ThisWindow.Run PROCEDURE(USHORT Number,BYTE Request)

ReturnValue          BYTE,AUTO

  CODE
  ReturnValue = PARENT.Run(Number,Request)
  IF SELF.Request = ViewRecord
    ReturnValue = RequestCancelled                         ! Always return RequestCancelled if the form was opened in ViewRecord mode
  ELSE
    GlobalRequest = Request
    UpdateProjects
    ReturnValue = GlobalResponse
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
    OF ?StartBtn
      ThisWindow.Update()
      MyState = MyState:Start
      Window{PROP:Timer} = 5      
    OF ?StopBtn
      ThisWindow.Update()
      MyTrace = -1
      Window{PROP:Timer} = 0 
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
    ThisWebClient.TakeEvent()                 ! Generated by NetTalk Extension
  ReturnValue = PARENT.TakeEvent()
    Case EVENT()
    Of EVENT:Timer
        Case MyState
        Of MyState:Start
            MyTrace.Trace('[Main][ThisWindow.TakeEvent][EVENT:Timer][MyState:Start]')
            FileNdx = 0
            MyState = 0
            Free(LinksQ)
            !Window{PROP:Timer} = 0
            ThisWebClient.Fetch(Pro:MainUrl)
        !Of MyState:Split
        Of MyState:GetPages
            MyTrace.Trace('[Main][ThisWindow.TakeEvent][EVENT:Timer][MyState:GetPages]')
  
            If Not Records(LinksQ)
                MyState = MyState:ParsePage
            End
  
            TempSt.SetValue(ThisWebClient.ThisPage.GetValue())
            TempSt.SplitBetween('href="','"',True,True)
            Loop x = 1 to TempSt.Records()
                LneSt.SetValue(TempSt.GetLine(x))
                !If Not LneSt.FindChars('http')
                !    MenuSt.AppendA(Clip(Pro:MainUrl) & TempSt.GetLine(x),True,'<13,10>')
                !    LneSt.SetValue(Clip(Pro:MainUrl) & TempSt.GetLine(x))
                !ELSE
                !    MenuSt.AppendA(TempSt.GetLine(x),True,'<13,10>')
                !    LneSt.SetValue(TempSt.GetLine(x))
                !End
                If not LneSt.Instring('http') And not LneSt.Instring(Clip(Pro:MainUrl))
                    If LneSt.Slice(1,1) = '/'
                        LneSt.Prepend(Clip(Pro:MainUrl))
                    ELSE
                        LneSt.Prepend(Clip(Pro:MainUrl) & '/')
                    End
                End
                !LneSt.Remove(LneSt.UrlParametersOnly(LneSt.GetValue()))
                If LneSt.ExtensionOnly(LneSt.UrlFileOnly())
                    LinksQ.Link = LneSt.GetValue()
                    LinksQ.FileName = FileSt.UrlFileOnly(LinksQ.Link) !LneSt.UrlFileOnly(LneSt.GetValue())
                    LinksQ.FileExt = LneSt.ExtensionOnly(LinksQ.FileName)   !LneSt.ExtensionOnly(LneSt.UrlFileOnly(LneSt.GetValue()))   !
                    If LinksQ.FileExt
                        Get(LinksQ,LinksQ.Link)
                        If Errorcode()
                            Add(LinksQ)
                        End
                    End
                ELSE
                    If LneSt.GetValue() <> Clip(Pro:MainUrl) & '/' And LneSt.Instring(Clip(Pro:MainUrl))
                        If Not LneSt.Instring('anchor')
                            PagesQ.PageLink = LneSt.UrlProtocolOnly() & '://' & LneSt.UrlHostOnly() & '/' & LneSt.UrlPathOnly() !LneSt.GetValue()
                            Get(PagesQ,PagesQ.PageLink)
                            If Errorcode()
                                Add(PagesQ)
                            End
                        End
                    End
                End
  
  
            End
        Of MyState:GetPage
            MyState = MyState:WaitOnPage
  
            If Not Records(PagesQ)
                MyState = 0
            End
            Get(PagesQ,1) 
  
            MyTrace.Trace('[Main][ThisWindow.TakeEvent][EVENT:Timer][MyState:GetPage]<9>PageLink[' & Clip(PagesQ.PageLink) & ']')
  
            ThisWebClient.Fetch(PagesQ.PageLink)
        Of MyState:WaitOnPage
            MyTrace.Trace('[Main][ThisWindow.TakeEvent][EVENT:Timer][MyState:WaitOnPage]')  
  
        Of MyState:ParsePage
            MyTrace.Trace('[Main][ThisWindow.TakeEvent][EVENT:Timer][MyState:ParsePage]') 
  
            ThisWebClient.ThisPage.SetValue(FullPageSt.GetValue())
  
            If Records(PagesQ)
                MyState = MyState:GetPage
            ELSE
                !Window{PROP:Timer} = 0
                MyState = MyState:GetFile
            End
            
            ThisWebClient.ThisPage.RemoveAttributes('li')
            !self.ThisPage.FormatHTML()
            TempSt.SetValue(ThisWebClient.ThisPage.GetValue())
            TempSt.SplitBetween('src="','"',True,True)
            Loop x = 1 to TempSt.Records()
                !SrcSt.AppendA(TempSt.GetLine(x),True,'<13,10>')
                LneSt.SetValue(TempSt.GetLine(x))
                !If not LneSt.Instring('http')
                !    LneSt.Prepend('http:')
                !End
                If not LneSt.Instring('http') And not LneSt.Instring(Clip(Pro:MainUrl))
                    If LneSt.Slice(1,1) = '/'
                        LneSt.Prepend(Clip(Pro:MainUrl))
                    ELSE
                        LneSt.Prepend(Clip(Pro:MainUrl) & '/')
                    End
                End
  
                LinksQ.Link = LneSt.UrlProtocolOnly() & '://' & LneSt.UrlHostOnly() & '/' & LneSt.UrlPathOnly() !LneSt.GetValue() !TempSt.GetLine(x)
  
                FileSt.SetValue(LinksQ.Link,True)
  
                FileSt.Split('<32>')
                !MyTrace.Trace('')
                !Loop x = 1 to FileSt.Records()
                !    MyTrace.Trace('<9>' & x & '<9>' & FileSt.GetLine(x))
                !END
                LinksQ.Link = FileSt.GetLine(1)
  
                LinksQ.FileName = FileSt.UrlFileOnly(LinksQ.Link) !LneSt.UrlFileOnly(LneSt.GetValue())
                LinksQ.FileExt = LneSt.ExtensionOnly(LinksQ.FileName)   !LneSt.ExtensionOnly(LneSt.UrlFileOnly(LneSt.GetValue()))   !
                !LneSt.U
                If LinksQ.FileExt
                    Get(LinksQ,LinksQ.Link)
                    If Errorcode()
                        Add(LinksQ)
                    End
                End
            End
            TempSt.SetValue(ThisWebClient.ThisPage.GetValue())
            TempSt.SplitBetween('srcset="','"',True,True)
            Loop x = 1 to TempSt.Records()
                !SrcSetSt.AppendA(TempSt.GetLine(x),True,'<13,10>')
                LneSt.SetValue(TempSt.GetLine(x))
                !If not LneSt.Instring('http')
                !    LneSt.Prepend('http:')
                !End
                If not LneSt.Instring('http') And not LneSt.Instring(Clip(Pro:MainUrl))
                    If LneSt.Slice(1,1) = '/'
                        LneSt.Prepend(Clip(Pro:MainUrl))
                    ELSE
                        LneSt.Prepend(Clip(Pro:MainUrl) & '/')
                    End
                End
  
                LinksQ.Link = LneSt.UrlProtocolOnly() & '://' & LneSt.UrlHostOnly() & '/' & LneSt.UrlPathOnly() !LneSt.GetValue() !TempSt.GetLine(x)
  
                FileSt.SetValue(LinksQ.Link,True)
  
                FileSt.Split('<32>')
                !MyTrace.Trace('')
                !Loop x = 1 to FileSt.Records()
                !    MyTrace.Trace('<9>' & x & '<9>' & FileSt.GetLine(x))
                !END
                LinksQ.Link = FileSt.GetLine(1)
                LinksQ.FileName = FileSt.UrlFileOnly(LinksQ.Link) !LneSt.UrlFileOnly(LneSt.GetValue())
                LinksQ.FileExt = LneSt.ExtensionOnly(LinksQ.FileName) !LneSt.ExtensionOnly(LneSt.UrlFileOnly(LneSt.GetValue()))   !
  
                If LinksQ.FileExt
                    Get(LinksQ,LinksQ.Link)
                    If Errorcode()
                        Add(LinksQ)
                    End
                End
            End
        Of MyState:GetFile
            If Records(LinksQ)
                MyState = MyState:WaitOnFile
                Get(LinksQ,1)
                
                MyTrace.Trace('[Main][ThisWindow.TakeEvent][EVENT:Timer][MyState:GetFile]<9>Link[' & Clip(LinksQ.Link) & ']')
  
                If LinksQ.FileExt
                    MyState = MyState:WaitOnFile
                    ThisWebClient.Fetch(LinksQ.Link)
                ELSE
                    Delete(LinksQ)
                    MyState = MyState:GetFile
                End
            ELSE
                MyState = -1
                Window{PROP:Timer} = 0
            End
        Of MyState:WaitOnFile
            MyTrace.Trace('[Main][ThisWindow.TakeEvent][EVENT:Timer][MyState:WaitOnFile]') 
        End
    END
  
  !MyState             Long
  !MyState:Start       Equate(1)  
  
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
        post(event:visibleondesktop)
    END
    RETURN ReturnValue
  END
  ReturnValue = Level:Fatal
  RETURN ReturnValue

local.RemoveSection           Procedure(StringTheory pSt,string pTag,long pExclusive)
RemoveSt        StringTheory
LneSt           StringTheory
DelimiterSt     StringTheory
WordSt          StringTheory
StartPos        Long
EndPos          Long
DelimStartPos   Long
DelimEndPos     Long
Left_           String(10)
Right           String(10)
x               Long
    code
    RemoveSt.SetValue(pSt.GetValue())
    DelimiterSt.SetValue(pTag,True)
    StartPos = 1
    MyTrace.Trace('')
    LOOP
        x+=1
        StartPos = RemoveSt.WordStart(StartPos, 0, st:Forwards,'<32>')
        if StartPos < 1 then break.
        EndPos = RemoveSt.WordEnd(StartPos, 0, '<32>',True)

        WordSt.SetValue(RemoveSt.Between('','',StartPos,EndPos,True,False))
        WordSt.SetValue(WordSt.Left())
        WordSt.Trim()

        MyTrace.Trace(WordSt.GetValue())

        LneSt.AppendA(WordSt.GetValue(),False,'<32>')

        

        If StartPos > DelimiterSt.Length()
            If WordSt.FindChars(DelimiterSt.GetValue(),2) !WordSt.GetValue() = DelimiterSt.GetValue()
                EndPos = LneSt.FindChars(DelimiterSt.GetValue(),2) + (DelimiterSt.Length())
                LneSt.RemoveFromPosition(EndPos,(LneSt.Length()-EndPos)+1)
                If pExclusive = 1
                    DelimStartPos = (LneSt.Length()-DelimiterSt.Length()) + 1
                    DelimEndPos = LneSt.Length()
                    LneSt.RemoveFromPosition(DelimStartPos,(DelimEndPos-DelimStartPos)+1)
                End
                EndPos = (StartPos + LneSt.Length()) - 1
                Break
            End
        End

        if EndPos < StartPos or EndPos > RemoveSt.Length()
            break
        End
        StartPos = EndPos + 1

    End
    MyTrace.Trace('')
    !MyTrace.Trace(LneSt.GetValue())
    !MyTrace.Trace('')
    pSt.Remove(LneSt.GetValue())









!    Left_ = clip(pTag) !'<' & clip(pTag) & '<32>'
!    Right = clip(pTag) !'</' & clip(pTag) & '>'
!    StartPos = RemoveSt.FindNth(Clip(Left_),1) !RemoveSt.FindChars(Clip(Left_))
!    MyTrace.Trace('')
!    loop     
!        EndPos = RemoveSt.FindChars(Clip(Right),StartPos+1) !+ Size(Right)      
!        LneSt.SetValue(RemoveSt.FindBetween(Clip(Left_), Clip(Right), StartPos, EndPos,True,False))     
!        if StartPos = 0         
!            break     
!        else         
!            ! do something with the returned betweenVal
!            MyTrace.Trace(LneSt.GetValue())   
!            !RemoveSt.Remove(LneSt.GetValue())  
!        end   
!        ! Reset pStart for next iteration. If not doing pExclusive then pStart = pEnd + 1
!        StartPos = EndPos + 1 !len(pRight) + 1
!    end
!    MyTrace.Trace('')    
!    pSt.SetValue(RemoveSt.GetValue())
    !Return RemoveSt.GetValue()
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
  SELF.SetStrategy(?Close,100,100,0,0)
  SELF.SetStrategy(?SHEET1,0,0,100,100)
  SELF.SetStrategy(?ListPageLinks,0,0,100,100)
  SELF.SetStrategy(?ListFileLinks,0,0,100,100)

BrwProjects.Init PROCEDURE(SIGNED ListBox,*STRING Posit,VIEW V,QUEUE Q,RelationManager RM,WindowManager WM)

  CODE
  PARENT.Init(ListBox,Posit,V,Q,RM,WM)
  IF WM.Request <> ViewRecord                              ! If called for anything other than ViewMode, make the insert, change & delete controls available
    SELF.InsertControl=?Insert
    SELF.ChangeControl=?Change
    SELF.DeleteControl=?Delete
  END


ThisWebClient.ErrorTrap PROCEDURE(string errorStr,string functionName)


  CODE
    If Records(PagesQ)
        Delete(PagesQ)
        MyState = MyState:GetPage
    ELSE
        If Records(LinksQ)
            Delete(LinksQ)
            MyState = MyState:GetFile
        End
    End    
  PARENT.ErrorTrap(errorStr,functionName)


ThisWebClient.PageReceived PROCEDURE

i           Long
LinkSt      StringTheory

  CODE
    FileNdx += 1
    If self.ThisPage.Length()
        If self.ServerResponse = 200
            FullPageSt.SetValue(self.ThisPage.GetValue())
            !self.RemoveHeader()
            Case MyState
            Of 0
                MyState = MyState:GetPages
                If Not EXISTS(Clip(Pro:SaveTo) & '\Pages')
                    ds_CreateDirectory(Clip(Pro:SaveTo) & '\Pages')
                End
                self.RemoveHeader()
                self.ThisPage.SaveFile(Clip(Pro:SaveTo) & '\Pages\Index.html')
            Of MyState:WaitOnPage
                !MyTrace.Trace('')
                !MyTrace.Trace(self.ThisPage.GetValue())
  
                LinkSt.SetValue(PagesQ.PageLink,True)
                LinkSt.Split('/')
                !MyTrace.Trace('')
                !Loop i = 1 to LinkSt.Records()
                !    MyTrace.Trace('<9>' & i & '<9>' & LinkSt.GetLine(i))
                !End
                self.RemoveHeader()
                If Pro:SaveTo
                    If Not EXISTS(Clip(Pro:SaveTo) & '\Pages')
                        ds_CreateDirectory(Clip(Pro:SaveTo) & '\Pages')
                    End
                    self.ThisPage.SaveFile(Clip(Pro:SaveTo) & '\Pages\' & LinkSt.GetLine(LinkSt.Records()) & '.html')
                ELSE
                    self.ThisPage.SaveFile('.\Extracted Files\Pages\' & LinkSt.GetLine(LinkSt.Records()) & '.html')
                End
  
  
                !MyState = MyState:GetPage
                Delete(PagesQ)
                If Records(PagesQ)
                    MyState = MyState:ParsePage
                ELSE
                    MyState = MyState:GetFile
                End
            Of MyState:WaitOnFile
                self.RemoveHeader()
                If Records(LinksQ)
                    Delete(LinksQ)
                    MyState = MyState:GetFile
                    Case Upper(Clip(LinksQ.FileExt))
                    Of 'HTML'
                    OrOf 'HTM'
                        If Pro:SaveTo
                            If Not EXISTS(Clip(Pro:SaveTo) & '\Pages')
                                ds_CreateDirectory(Clip(Pro:SaveTo) & '\Pages')
                            End
                            self.ThisPage.SaveFile(Clip(Pro:SaveTo) & '\Pages\' & Clip(LinksQ.FileName))
                        End
                    Of 'JS'
                        If Pro:SaveTo
                            If Not EXISTS(Clip(Pro:SaveTo) & '\JavaScript')
                                ds_CreateDirectory(Clip(Pro:SaveTo) & '\JavaScript')
                            End
                            self.ThisPage.SaveFile(Clip(Pro:SaveTo) & '\JavaScript\' & Clip(LinksQ.FileName))
                        End
                    Of 'CSS'
                        If Pro:SaveTo
                            If Not EXISTS(Clip(Pro:SaveTo) & '\CSS')
                                ds_CreateDirectory(Clip(Pro:SaveTo) & '\CSS')
                            End
                            self.ThisPage.SaveFile(Clip(Pro:SaveTo) & '\CSS\' & Clip(LinksQ.FileName))
                        End
                    Of 'JPG'
                    OrOf 'JPEG'
                    OrOf 'PNG'
                    OrOf 'WEBP'
                    OrOf 'GIF'
                        If Pro:SaveTo
                            If Not EXISTS(Clip(Pro:SaveTo) & '\Images')
                                ds_CreateDirectory(Clip(Pro:SaveTo) & '\Images')
                            End
                            self.ThisPage.SaveFile(Clip(Pro:SaveTo) & '\Images\' & Clip(LinksQ.FileName))
                        End
                    Of 'GZ'
                        If Pro:SaveTo
                            If Not EXISTS(Clip(Pro:SaveTo) & '\Zipped')
                                ds_CreateDirectory(Clip(Pro:SaveTo) & '\Zipped')
                            End
                            self.ThisPage.SaveFile(Clip(Pro:SaveTo) & '\Zipped\' & Clip(LinksQ.FileName))
                        End
                    Of 'ASPX'
                        If Pro:SaveTo
                            If Not EXISTS(Clip(Pro:SaveTo) & '\ASPX')
                                ds_CreateDirectory(Clip(Pro:SaveTo) & '\ASPX')
                            End
                            self.ThisPage.SaveFile(Clip(Pro:SaveTo) & '\ASPX\' & Clip(LinksQ.FileName))
                        End
                    Of 'SVG'
                        If Pro:SaveTo
                            If Not EXISTS(Clip(Pro:SaveTo) & '\SVG')
                                ds_CreateDirectory(Clip(Pro:SaveTo) & '\SVG')
                            End
                            self.ThisPage.SaveFile(Clip(Pro:SaveTo) & '\SVG\' & Clip(LinksQ.FileName))
                        End
                    OF 'JSON'
                        If Pro:SaveTo
                            If Not EXISTS(Clip(Pro:SaveTo) & '\JSON')
                                ds_CreateDirectory(Clip(Pro:SaveTo) & '\JSON')
                            End
                            self.ThisPage.SaveFile(Clip(Pro:SaveTo) & '\JSON\' & Clip(LinksQ.FileName))
                        End
                    Of 'ICO'
                        If Pro:SaveTo
                            If Not EXISTS(Clip(Pro:SaveTo) & '\Icons')
                                ds_CreateDirectory(Clip(Pro:SaveTo) & '\Icons')
                            End
                            self.ThisPage.SaveFile(Clip(Pro:SaveTo) & '\Icons\' & Clip(LinksQ.FileName))
                        End
                    Of 'XML'
                        If Pro:SaveTo
                            If Not EXISTS(Clip(Pro:SaveTo) & '\XML')
                                ds_CreateDirectory(Clip(Pro:SaveTo) & '\XML')
                            End
                            self.ThisPage.SaveFile(Clip(Pro:SaveTo) & '\XML\' & Clip(LinksQ.FileName))
                        End
                    Of 'RSS'
                        If Pro:SaveTo
                            If Not EXISTS(Clip(Pro:SaveTo) & '\RSS')
                                ds_CreateDirectory(Clip(Pro:SaveTo) & '\RSS')
                            End
                            self.ThisPage.SaveFile(Clip(Pro:SaveTo) & '\RSS\' & Clip(LinksQ.FileName))
                        End
                    ELSE
                        If Pro:SaveTo
                            If Not EXISTS(Clip(Pro:SaveTo) & '\Other')
                                ds_CreateDirectory(Clip(Pro:SaveTo) & '\Other')
                            End
                            self.ThisPage.SaveFile(Clip(Pro:SaveTo) & '\Other\' & Clip(LinksQ.FileName))
                        End
                    End
                    !self.ThisPage.SaveFile('.\Extracted Files\' & Clip(LinksQ.FileName))
                ELSE
                    MyState = -1
                    Window{PROP:Timer} = 0
                End
            End
        ELSE
            Delete(PagesQ)
            If Records(PagesQ)
                MyState = MyState:GetPage
            ELSE
                If Records(LinksQ)
                    Delete(LinksQ)
                    MyState = MyState:GetFile
                End
            End
  
        End
    Else
        Delete(PagesQ)
        If Records(PagesQ)
            MyState = MyState:GetPage
        ELSE
            If Records(LinksQ)
                Delete(LinksQ)
                MyState = MyState:GetFile
            End
        End  
    End  
  PARENT.PageReceived

