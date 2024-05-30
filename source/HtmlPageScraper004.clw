

   MEMBER('HtmlPageScraper.clw')                           ! This is a MEMBER module

                     MAP
                       INCLUDE('HTMLPAGESCRAPER004.INC'),ONCE        !Local module procedure declarations
                     END


!!! <summary>
!!! Generated from procedure template - Source
!!! </summary>
ds_Stop              PROCEDURE  (<string StopText>)        ! Declare Procedure
returned              long
TempTimeOut           long(0)
TempNoLogging         long(0)
Loc:StopText    string(4096)

  CODE
  if omitted(1) or StopText = '' then
    if ThisMessageBox.GetGlobalSetting('TranslationFile') <> ''
      Loc:StopText = getini('MessageBox_Text','StopDefault','Exit?',ThisMessageBox.GetGlobalSetting('TranslationFile'))
    elsif ThisMessageBox.GetGlobalSetting('StopDefault') <> ''
      Loc:StopText = ThisMessageBox.GetGlobalSetting('StopDefault')
    else
      Loc:StopText = 'Exit?'
    end
  else
    Loc:StopText = StopText
  end
  if ThisMessageBox.GetGlobalSetting('TimeOut') > 0
    TempTimeOut = ThisMessageBox.GetGlobalSetting('TimeOut')
    ThisMessageBox.SetGlobalSetting('TimeOut', 0)
  end
  if ThisMessageBox.GetGlobalSetting('TranslationFile') <> ''
    returned = ds_Message(Loc:StopText,getini('MessageBox_Text','StopHeader','Stop',ThisMessageBox.GetGlobalSetting('TranslationFile')),ICON:Hand,BUTTON:Abort+BUTTON:Ignore,BUTTON:Abort)
  elsif ThisMessageBox.GetGlobalSetting('StopHeader') <> ''
    returned = ds_Message(Loc:StopText,ThisMessageBox.GetGlobalSetting('StopHeader'),ICON:Hand,BUTTON:Abort+BUTTON:Ignore,BUTTON:Abort)
  else
    returned = ds_Message(Loc:StopText,'Stop',ICON:Hand,BUTTON:Abort+BUTTON:Ignore,BUTTON:Abort)
  end
  if returned = BUTTON:Abort then halt .
  if TempTimeOut
    ThisMessageBox.SetGlobalSetting('TimeOut',TempTimeOut)
  end
