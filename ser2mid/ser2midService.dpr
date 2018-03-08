program ser2midService;
uses
  SvcMgr,
  ServiceModule in 'ServiceModule.pas' {AppService: TUniGUIService},
  Forms,
  Windows,
  SysUtils,
  Classes,
  ServerModule in 'ServerModule.pas' {FServerModule: TUniGUIServerModule},
  utils in 'utils.pas';

{$R *.res}

begin

    APP_NAME := 'ser2midService';
    APP_VERSION := APP_NAME + ' ' + FileVersion( paramStr(0) );

    if upperCase( paramStr(1)  ) = '/APP' then
       with Forms.Application do begin
           Initialize;
           Title := APP_VERSION;
           CreateForm( TFServerModule, FServerModule );
           Run;
       end
    else
        with SvcMgr.Application do begin
            //al ser servicio windows lo cambia a system32...
           SetCurrentDir( extractFileDir(paramStr(0)) );
           if not DelayInitialize or Installing then
               Initialize;
           Title := APP_VERSION;
           CreateForm( TFServerModule, FServerModule );
           CreateForm( TAppService, AppService);
           Run;
       end;


end.


