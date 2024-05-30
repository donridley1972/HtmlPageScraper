  MEMBER('HtmlPageScraper.clw')

  INCLUDE('ABFILE.INC'),ONCE

  MAP
HtmlPageScraper_BC0:DctInit    PROCEDURE()
HtmlPageScraper_BC0:DctKill    PROCEDURE()
HtmlPageScraper_BC0:FilesInit  PROCEDURE()
  END

Hide:Access:Projects CLASS(FileManager),TYPE               ! FileManager for Projects
Init                   PROCEDURE(),DERIVED
Kill                   PROCEDURE(),DERIVED
PrimeFields            PROCEDURE(),PROC,DERIVED
UseFile                PROCEDURE(BYTE UseType = UseType:Uses),BYTE,PROC,DERIVED
                     END


Hide:Relate:Projects CLASS(RelationManager),TYPE           ! RelationManager for Projects
Init                   PROCEDURE
Kill                   PROCEDURE(),DERIVED
                     END


_Hide:Access:Projects &Hide:Access:Projects,AUTO,THREAD
_Hide:Relate:Projects &Hide:Relate:Projects,AUTO,THREAD


HtmlPageScraper_BC0:DctInit PROCEDURE()
  CODE
  _Hide:Access:Projects &= NEW(Hide:Access:Projects)
  _Hide:Relate:Projects &= NEW(Hide:Relate:Projects)
  Relate:Projects &= _Hide:Relate:Projects


HtmlPageScraper_BC0:FilesInit PROCEDURE()
  CODE
  _Hide:Relate:Projects.Init


HtmlPageScraper_BC0:DctKill PROCEDURE()
  CODE
  _Hide:Relate:Projects.Kill
  DISPOSE(_Hide:Relate:Projects)


Hide:Relate:Projects.Init PROCEDURE()
  CODE
  _Hide:Access:Projects.Init
  SELF.Init(Access:Projects,1)


Hide:Access:Projects.Init PROCEDURE

  CODE
  SELF.Initialized = False
  SELF.Buffer &= Pro:Record
  SELF.FileNameValue = 'Projects'
  SELF.SetErrors(GlobalErrors)
  SELF.File &= Projects
  PARENT.Init
  Access:Projects &= SELF
  SELF.DoNotAssertOnCloseMisuse = 1  ! ABCDefaults
   IF SELF.File{prop:sqldriver} THEN SELF.create = 0.         ! ABCDefaults


Hide:Access:Projects.Kill PROCEDURE

  CODE
  PARENT.Kill
  Access:Projects &= NULL                                  ! File manager has been destroyed to assign null, it is an error to reference this after this point


Hide:Access:Projects.PrimeFields PROCEDURE

  CODE
  Pro:Guid = Glo:st.MakeGuid()                             ! Assign initial field value
  PARENT.PrimeFields


Hide:Access:Projects.UseFile PROCEDURE(BYTE UseType = UseType:Uses)

ReturnValue          BYTE,AUTO

  CODE
  IF UseType ~= UseType:Initialize
    SELF.UseFile(UseType:Initialize)                       !Recursive call to ensure initialization takes place
  END
  IF UseType = UseType:Initialize AND ~SELF.Initialized
    SELF.InUseFile = True
    SELF.Init(Projects,GlobalErrors)
    SELF.Create = 1
    SELF.LockRecover = 10
    SELF.AddKey(Pro:PKProGuidKey,'Pro:PKProGuidKey',0)
    SELF.AddKey(Pro:ProDescriptionKey,'Pro:ProDescriptionKey',0)
    SELF.AddKey(Pro:ProMainUrlKey,'Pro:ProMainUrlKey',0)
  END
                                                           !SELF.InUseFile will be set to False in PARENT
  IF SELF.File{prop:sqldriver} THEN SELF.create = 0.          ! ABCDefaults
  ReturnValue = PARENT.UseFile(UseType)
  RETURN ReturnValue


Hide:Relate:Projects.Kill PROCEDURE

  CODE
  _Hide:Access:Projects.Kill                               ! Kill the file manager
  PARENT.Kill
  Relate:Projects &= NULL                                  ! Assign NULL to the RelationManager's FileManager reference, it is an error to reference this after this point
  DISPOSE(_Hide:Access:Projects)                           ! destroy the file manager

