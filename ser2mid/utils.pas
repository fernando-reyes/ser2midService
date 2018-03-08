unit utils;

interface
uses
    Inifiles, Classes;

function FileVersion(const FileName: String): String;
function str2Hex( binStr:AnsiString ; size:Integer ):AnsiString;
function hex2Bin(hex:AnsiString):AnsiString;

const
        _MSG_ERROR_  :Integer = 1; _MSG_ALWAYS_ :Integer = 1;
        _MSG_WARNING_:Integer = 2;
        _MSG_HINT_   :Integer = 3;
        _MSG_DEBUG_  :Integer = 4;
        CR   = #$0D;
        LF   = #$0A;
        CRLF = CR+LF;

var APP_VERSION:String;
    APP_NAME:String;

implementation
uses
    Windows,
    sysUtils;


function FileVersion(const FileName: String): String;
var
  VerInfoSize: Cardinal;
  VerValueSize: Cardinal;
  Dummy: Cardinal;
  PVerInfo: Pointer;
  PVerValue: PVSFixedFileInfo;
begin
  Result := '';
  VerInfoSize := GetFileVersionInfoSize(PChar(FileName), Dummy);
  GetMem(PVerInfo, VerInfoSize);
  try
    if GetFileVersionInfo(PChar(FileName), 0, VerInfoSize, PVerInfo) then
      if VerQueryValue(PVerInfo, '\', Pointer(PVerValue), VerValueSize) then
        with PVerValue^ do
          Result := Format('v%d.%d.%d.%d', [
            HiWord(dwFileVersionMS), //Major
            LoWord(dwFileVersionMS), //Minor
            HiWord(dwFileVersionLS), //Release
            LoWord(dwFileVersionLS)]); //Build
  finally
    FreeMem(PVerInfo, VerInfoSize);
  end;
end;

function str2Hex( binStr:AnsiString ; size:Integer ):AnsiString;
     var i:integer;
         hexStr:AnsiString;
   begin
        result := '';
        for i := 1 to size do
            result := result + intToHex( byte(binStr[i]) , 2 );
     end;

function hex2Bin(hex:AnsiString):AnsiString;
     var i:integer;
   begin
         result := '';
         for i:=1 to length(hex) div 2 do
             result := result + AnsiChar(chr( strToInt( '$'+hex[i*2-1]+hex[i*2] ) ) );
    end;


end.

