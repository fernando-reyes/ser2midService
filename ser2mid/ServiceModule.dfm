object AppService: TAppService
  OldCreateOrder = False
  OnCreate = UniGUIServiceCreate
  DisplayName = 'ser2midService'
  ServiceStartName = 'NT AUTHORITY\SYSTEM'
  AfterInstall = UniGUIServiceAfterInstall
  Left = 598
  Top = 420
  Height = 150
  Width = 215
end
