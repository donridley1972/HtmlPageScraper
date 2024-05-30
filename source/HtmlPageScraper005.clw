

   MEMBER('HtmlPageScraper.clw')                           ! This is a MEMBER module

                     MAP
                       INCLUDE('HTMLPAGESCRAPER005.INC'),ONCE        !Local module procedure declarations
                     END


!!! <summary>
!!! Generated from procedure template - Source
!!! </summary>
ds_Halt              PROCEDURE  (UNSIGNED Level=0,<STRING HaltText>) ! Declare Procedure
TempNoLogging         long(0)

  CODE
  if ~omitted(2) then
    if ThisMessageBox.GetGlobalSetting('TranslationFile') <> ''
      ds_Message(HaltText,getini('MessageBox_Text','HaltHeader','Halt',ThisMessageBox.GetGlobalSetting('TranslationFile')),ICON:Hand)
    elsif ThisMessageBox.GetGlobalSetting('HaltHeader') <> ''
      ds_Message(HaltText,ThisMessageBox.GetGlobalSetting('HaltHeader'),ICON:Hand)
    else
      ds_Message(HaltText,'Halt',ICON:Hand)
    end
  end
  system{prop:HaltHook} = 0
  HALT(Level)
