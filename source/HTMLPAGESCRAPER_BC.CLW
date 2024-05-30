  MEMBER('HtmlPageScraper.clw')

  PRAGMA('define(init_priority=>3)')

  MAP
    MODULE('HtmlPageScraper_BC0.CLW')
HtmlPageScraper_BC0:DctInit             PROCEDURE()
HtmlPageScraper_BC0:DctKill             PROCEDURE()
HtmlPageScraper_BC0:FilesInit           PROCEDURE()
    END
  END

DctInit PROCEDURE()
  CODE
  HtmlPageScraper_BC0:DctInit
  HtmlPageScraper_BC0:FilesInit


DctKill PROCEDURE()
  CODE
  HtmlPageScraper_BC0:DctKill

