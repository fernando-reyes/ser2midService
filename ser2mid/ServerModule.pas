unit ServerModule;

interface

uses
  Classes, CPort, Vcl.ExtCtrls,WinApi.Messages,Windows,Forms, Vcl.Menus;

type
  TFServerModule = class(TForm)
    ComPort1: TComPort;
    TrayIcon1: TTrayIcon;
    PopupMenu1: TPopupMenu;
    Cerrar1: TMenuItem;
    procedure FServerModuleCreate(Sender: TObject);
    procedure FServerModuleDestroy(Sender: TObject);
    procedure ComPort1AfterClose(Sender: TObject);
    procedure ComPort1AfterOpen(Sender: TObject);
    procedure ComPort1Exception(Sender: TObject; TComException: TComExceptions;
      ComportMessage: string; WinError: Int64; WinMessage: string);
    procedure Cerrar1Click(Sender: TObject);
    procedure ComPort1RxChar(Sender: TObject; Count: Integer);
  protected
  private
    logFile:TFileStream;
    debug_level:Integer;
    lastMsg:String;
    bufStr:String;
    bufLen:Integer;


    procedure WMDEVICECHANGE(var Msg: TMessage); message WM_DEVICECHANGE;

  public
    userList:TStringList;
    procedure writeLog( x:AnsiString );
    procedure writeLogLN( const x:AnsiString );overload;
    procedure writeLogLN( const x:AnsiString ; level:Integer );overload;

  end;


var
    FServerModule:TFServerModule;

implementation

{$R *.dfm}

uses
    sysUtils,
    Math,
    utils,
    IniFiles,
    StrUtils,
//    IdGlobal,
    MMSystem;

var
    FWinHandle : HWND;
    mo: HMIDIOUT;
    MIDI_DEVICE:word;

type
    PDevBroadcastHdr  = ^DEV_BROADCAST_HDR;
    DEV_BROADCAST_HDR = packed record
        dbch_size       : DWORD;
        dbch_devicetype : DWORD;
        dbch_reserved   : DWORD;
    end;

    PDev_Broadcast_Port = ^DEV_BROADCAST_PORT;
    DEV_BROADCAST_PORT = record
        dbcp_size:DWORD ;
        dbcp_devicetype:DWORD ;
        dbcp_reserved:DWORD ;
        dbcp_name:array[0..0] of AnsiChar ;
    end;

procedure TFServerModule.WMDEVICECHANGE(var Msg: TMessage);
    const DBT_DEVICEARRIVAL         = $8000;    // system detected a new device
          DBT_DEVICEREMOVECOMPLETE  = $8004;    // device is gone
          DBT_DEVTYP_PORT           = $0003;    //lpt y com
      var devType: Integer;
          datos: PDevBroadcastHdr;
          port:string;
begin
    inherited;
    if ( Msg.WParam = DBT_DEVICEARRIVAL ) or ( Msg.WParam = DBT_DEVICEREMOVECOMPLETE ) then begin
        datos := PDevBroadcastHdr(Msg.lParam);
        if Datos^.dbch_devicetype = DBT_DEVTYP_PORT then
            if PChar(@PDev_Broadcast_Port(datos).dbcp_name) = upperCase( comPort1.Port ) then
                comPort1.connected := Msg.WParam = DBT_DEVICEARRIVAL;
    end;
end;


procedure TFServerModule.Cerrar1Click(Sender: TObject);
begin
    halt;
end;

procedure TFServerModule.ComPort1AfterClose(Sender: TObject);
begin
    writeLogLN( 'Desconectado...' );
end;

procedure TFServerModule.ComPort1AfterOpen(Sender: TObject);
begin
    writeLogLN( 'Conectado con ['+comPort1.Port+'/'+intToStr(comPort1.CustomBaudRate)+']' );
end;

procedure TFServerModule.ComPort1Exception(Sender: TObject; TComException: TComExceptions; ComportMessage: string; WinError: Int64; WinMessage: string);
      var newMsg:String;
begin
    newMsg := ComportMessage+'/'+WinMessage;
    if lastMsg<>newMsg then begin
        writeLogLN( 'Error:'+newMsg );
        lastMsg := newMsg;
    end;
end;

procedure TFServerModule.ComPort1RxChar(Sender: TObject; Count: Integer);
      var buf:String;
begin
    comport1.ReadStr(Buf,Count);
    bufStr := bufStr + buf;
    while length(bufStr)>=3 do begin
        buf := copy( bufStr , 1 ,3 );
        midiOutShortMsg(mo, ord(Buf[1]) + (ord(Buf[2]) shl 8) + (ord(Buf[3]) shl 16) );
        bufStr := copy( bufStr , 4 , length( bufStr ) );
    end;
end;

procedure TFServerModule.writeLog( x:AnsiString );
begin
    TThread.queue( nil,
        procedure
            begin
                x := formatDateTime( 'yyyy-mm-dd hh:nn:ss ' , now ) + x;
                logFile.writeBuffer( pointer(x)^ , length(x) );
             end );
end;

procedure TFServerModule.writeLogLN( const x:AnsiString );
begin
    writeLogLN( x ,  _MSG_ERROR_ );
end;

procedure TFServerModule.writeLogLN( const x:AnsiString ; level:Integer );
begin
    trayIcon1.Hint := x;
    if (level = _MSG_ERROR_ ) or (level <= debug_level ) then
        writeLog( x + CRLF );
end;


procedure TFServerModule.FServerModuleCreate(Sender: TObject);
      var logFileName:String;
          midiOutName:String;
          i:Integer;
          MIDIOUTCAPS_:MIDIOUTCAPS;
begin

    Application.ShowMainForm := False;

    MIDI_DEVICE := 0;

    with TMemIniFile.Create( ExtractFilePath( Application.ExeName ) + APP_NAME+'.conf') do begin

        comPort1.Port           :=              ReadString('midi'   ,'input_port'      ,'COM1');
        comPort1.CustomBaudRate := strToIntDef( ReadString('midi'   ,'input_bps'       ,'19200'), 19200 );
        midiOutName             :=              ReadString('midi'   ,'midi_output'    ,'Microsoft GS Wavetable Synth');

        debug_level                                 := strToIntDef( ReadString('global','debug_level','') , 0 );

        free;
    end;

    logFileName := ExtractFilePath( Application.ExeName ) + APP_NAME+'.log';
    if not fileExists( logFileName ) then
        fileClose( fileCreate( logFileName ) );
    logFile := TFileStream.create( logFileName , Math.ifThen(debug_level >= _MSG_ERROR_, fmCreate, fmOpenWrite ) or fmShareDenyWrite );
    logFile.Position := logFile.Size;

    writeLogLN( 'Sistema iniciado' , _MSG_ALWAYS_ );

    for i:=0 to midiOutGetNumDevs-1 do
        if midiOutGetDevCaps( i , @MIDIOUTCAPS_, sizeOf(MIDIOUTCAPS) ) = MMSYSERR_NOERROR then begin
            writeLogLN( 'Salida MIDI detectada ['+intToStr(i)+']:['+WideCharToString(MIDIOUTCAPS_.szPname)+']' , _MSG_ALWAYS_ );
            if midiOutName = WideCharToString(MIDIOUTCAPS_.szPname) then begin
                writeLogLN( 'Seleccionada:['+WideCharToString(MIDIOUTCAPS_.szPname)+']' , _MSG_ALWAYS_ );
                MIDI_DEVICE := i;
            end;
        end;


    FWinHandle := AllocateHWND(WMDEVICECHANGE) ;

    midiOutOpen(@mo, MIDI_DEVICE, 0, 0, CALLBACK_NULL);
    lastMsg := '@';
    bufLen:=0;
    bufStr:='';
    try
        comPort1.connected := true;
    except
        on e:Exception do
            writeLogLN( e.message );
    end;

end;

procedure TFServerModule.FServerModuleDestroy(Sender: TObject);
begin
    if comPort1.Connected then
        comPort1.Connected := false;
    writeLogLN( 'Sistema Finalizado' , _MSG_ALWAYS_ );
    DeallocateHWnd(FWinHandle);
    logFile.free;
end;

end.
