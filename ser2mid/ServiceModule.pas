unit ServiceModule;
//si hay problemas cambiar ServiceStartname : "NT AUTHORITY\LocalSystem" / "NT AUTHORITY\NetworkService"
interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, SvcMgr, Dialogs;

type
  TAppService = class(TService)
    procedure UniGUIServiceAfterInstall(Sender: TService);
    procedure UniGUIServiceCreate(Sender: TObject);
  private
  protected
  public
    function GetServiceController: TServiceController; override;
  end;

var
  AppService: TAppService;

implementation

{$R *.dfm}

uses
  Forms, Registry,Utils;

procedure ServiceController(CtrlCode: DWord); stdcall;
begin
  AppService.Controller(CtrlCode);
end;

function TAppService.GetServiceController: TServiceController;
begin
  Result := ServiceController;
end;

procedure TAppService.UniGUIServiceAfterInstall(Sender: TService);
      var Reg: TRegistry;
begin
    Reg := TRegistry.Create(KEY_READ or KEY_WRITE);
    try
        Reg.RootKey := HKEY_LOCAL_MACHINE;
        if Reg.OpenKey('\SYSTEM\CurrentControlSet\Services\' + Name, false) then begin
            Reg.WriteString('Description', APP_NAME );
            Reg.WriteInteger('DelayedAutostart',1);//delayed
            Reg.CloseKey;
        end;
    finally
        Reg.Free;
    end;
end;

procedure TAppService.UniGUIServiceCreate(Sender: TObject);
begin
    Name := APP_NAME+'Service';
    DisplayName := Name;
end;

end.

