program SendSMS;

{$APPTYPE CONSOLE}

uses
  SysUtils,
  Windows,
  Classes,
  System,
  ActiveX,
  ComObj;

const
  AUTHOR        = '0zon';
  INTERNAL_NAME = 'GCSMSSender';
  VERSION       = '0.1.2';

var
  Answer: String;

function IfThen(Value: Boolean; const True: string; False: string = ''): string;
begin
  if Value then
    Result := True
  else
    Result := False;
end;

function HTTPRequest(Method: String; Url: String; Data: String; RedirectMode: boolean; Header: TStringList; Proxy: string; var Error: Integer): String;
const
  SXH_PROXY_SET_PROXY = 2;
  HTTPREQUEST_SETCREDENTIALS_FOR_SERVER = 0;
  HTTPREQUEST_SETCREDENTIALS_FOR_PROXY = 1;

var
  ErrorCode, i: Integer;
  HttpReq: OLEVariant;
  Key, Value: string;
  ProxyServer, ProxyUsername, ProxyPassword: string;
  NewHeader: TStringList;
begin
  // Create the WinHttpRequest COM object
  HttpReq := CreateOLEObject('WinHttp.WinHttpRequest.5.1');

  // Initially set the return value of the function to ''
  Result := '';

  if (Proxy <> '') then begin
    if pos('@', Proxy) > 0 then begin
      ProxyServer := copy(Proxy, pos('@', Proxy) + 1, MaxInt);
      ProxyUsername := copy(Proxy, 1, pos('@', Proxy) - 1);
      if pos(':', ProxyUsername) > 0 then begin
        ProxyPassword := copy(ProxyUsername, pos(':', ProxyUsername) + 1, MaxInt);
        ProxyUsername := copy(ProxyUsername, 1, pos(':', ProxyUsername) - 1);
      end
      else
        ProxyPassword := '';
    end
    else begin
      ProxyServer := Proxy;
      ProxyUsername := '';
      ProxyPassword := '';
    end;

    //Set proxy server and bypass list
    ErrorCode := HttpReq.setProxy(SXH_PROXY_SET_PROXY,
      ProxyServer, '');
    if (ErrorCode <> S_OK) then begin
      Error := 1;
      Result := 'HTTP: Could not set Proxy server.';
      exit;
    end;

    //Set proxy username and password
    if (ProxyUsername <> '') then
    begin
      ErrorCode := HttpReq.SetCredentials(
        ProxyUsername, ProxyPassword,
        HTTPREQUEST_SETCREDENTIALS_FOR_PROXY);
      if (ErrorCode <> S_OK) then begin
        Error := 6;
        Result := 'HTTP: Could not call SetCredentials().';
        exit;
      end;
    end;
  end;

  ErrorCode := HttpReq.setAutoLogonPolicy(0);
  if (ErrorCode <> S_OK) then begin
    Error := 2;
    Result := 'HTTP: Could not call setAutoLogonPolicy.';
    exit;
  end;

  if (Method = 'GET') and (Data <> '') then
    ErrorCode := HttpReq.Open(Method, Url + '?' + Data, false)
  else
    ErrorCode := HttpReq.Open(Method, Url, false);
  if (ErrorCode <> S_OK) then begin
    Error := 4;
    Result := 'HTTP: Could not send GET request.';
    exit;
  end;

  HttpReq.SetRequestHeader('User-Agent', AUTHOR + '-' + INTERNAL_NAME + '-' + VERSION);

  // disable redirect
  HttpReq.Option(6) := RedirectMode;

  for i:= 0 to Header.Count - 1 do begin
    // Set HTTP Header
    Key := Header[i];
    Value := copy(Key, pos(': ', Key) + 2, MaxInt);
    Key := copy(Key, 1, pos(': ', Key) - 1);
    HttpReq.SetRequestHeader(Key, Value);
  end;

  if (Method = 'GET') or (Method = 'HEAD') then
    ErrorCode := HttpReq.Send()
  else begin
    if Header.Values['Content-Type'] = '' then
      HttpReq.SetRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
    HttpReq.SetRequestHeader('Content-Length', inttostr(Length(Data)));

    ErrorCode := HttpReq.Send(WideString(Data));
  end;

  if (ErrorCode <> S_OK) then begin
    Error := 7;
    Result := 'HTTP: Could not call Send().';
    exit;
  end;

  Result := HttpReq.ResponseText;
  {if HttpReq.Status div 100 = 2 then // 200, 201 ...
    Error := -1 // good response code
  else}
    Error := HttpReq.Status;

  if (HttpReq.Status div 100 = 3) and (HttpReq.Status <> 304) and (not HttpReq.Option(6)) then begin
    // manualy redirection
    NewHeader := TStringList.Create();
    NewHeader.Assign(Header);
    NewHeader.Add('Referer: ' + Url);
    Result := HTTPRequest({'GET'}Method, HttpReq.GetResponseHeader('Location'), Data, RedirectMode, NewHeader, Proxy, Error);
    NewHeader.Free;
    exit;
  end;
end;

function GetDateGC(dt: TDateTime): string;
var
  TimeZone: TTimeZoneInformation;
  Bias: Integer;
  ZoneID: Cardinal;
begin
  Bias := 0;
  ZoneID := GetTimeZoneInformation(TimeZone);
  case (ZoneID) of
   TIME_ZONE_ID_STANDARD:
     Bias := (-TimeZone.Bias - TimeZone.StandardBias);
   TIME_ZONE_ID_DAYLIGHT:
     Bias := (-TimeZone.Bias - TimeZone.DaylightBias);
   TIME_ZONE_ID_UNKNOWN:
     Bias := (-TimeZone.Bias);
  end;

  Result := FormatFloat('00.00', Bias / 60);
  if Bias >= 0 then
    Result := '+' + Result;
  Result[4] := ':';
  Result := FormatDateTime('yyyy-mm-dd_hh:nn:00', dt) + Result;
  Result[11] := 'T';
end;

function T(s: String): String;
begin
  Result := s;
  Result := StringReplace(StringReplace(Result, '>', '&gt;', [rfReplaceAll]), '<', '&lt;', [rfReplaceAll]);
end;

function IncMinutes(dt: TDateTime; m: Word): TDateTime;
begin
  Result := dt + m * 60 / 86400;
end;

function SendSMSviaGC(const Email, Passwd, Title, Content, Where, Proxy: String; var Answer: String): boolean;
const
  ReminderMinutes = 1;
  DeltaMinutes = 4; // must by greater than ReminderMinutes
  DurationMinutes = 5;
var
  url, pass, data, Auth: string;
  Header: TStringList;
  i, err: Integer;
  b: byte;
begin
  Result := false;

  // prepare password
  pass := passwd;
  if (Length(pass) > 3) and (pass[1] = '/') and (pass[2] <> '/') then begin //for examle: /21Q@RR
    b := StrToInt('$' + copy(pass, 2, 2));
    delete(pass, 1, 3);
    for i := 1 to Length(pass) do
      pass[i] := chr(ord(pass[i]) xor b);
  end;
  pass := StringReplace(pass, '//' , '/', [rfReplaceAll]);

  url := 'https://www.google.com/accounts/ClientLogin';
	data := 'Email=' + email + '&Passwd=' + pass + '&service=cl&source=' + AUTHOR + '-' + INTERNAL_NAME + '-' + VERSION;
  Header := TStringList.Create;
  Answer := HTTPRequest('POST', url, data, true, Header, Proxy, Err);
  Header.Text := Answer;
  if not (Err = 200) then begin
    if Header.Values['Error'] <> '' then
      Answer := 'GC: ' + Header.Values['Error']
    else
      Answer := 'HTTP: ' + IntToStr(Err) + ', ' + Answer;
    Header.Free;
    exit;
  end;

  Auth := Header.Values['Auth'];
  if Auth = '' then begin
    Header.Free;
    Answer := 'GC: ' + Header.Values['Error'];
    exit;
  end;

  url := 'http://www.google.com/calendar/feeds/default/private/full';
  data := '<atom:entry xmlns:atom="http://www.w3.org/2005/Atom">'+
  '  <atom:title type="text">' + T(Title) + '</atom:title>'+
  '  <atom:content type="text">' + T(Content) + '</atom:content>'+
  '  <gd:when xmlns:gd="http://schemas.google.com/g/2005" startTime="' + GetDateGC(IncMinutes(Now, DeltaMinutes)) + '" endTime="' + GetDateGC(IncMinutes(Now, DeltaMinutes + DurationMinutes)) + '">'+
  '    <gd:reminder minutes="' + inttostr(ReminderMinutes) + '" method="sms"/></gd:when>'+
  '  <gd:where xmlns:gd="http://schemas.google.com/g/2005" valueString="' + T(Where) + '"/></atom:entry>';
  Header.Clear;
  Header.NameValueSeparator := ':';
  Header.Add('authorization: GoogleLogin auth=' + Auth);
  Header.Add('Content-Type: application/atom+xml');
  Answer := HTTPRequest('POST', url, data, not true, Header, Proxy, Err);
  Header.Free;

  Result := (Err = 201); // HTTP response code 201: post created
end;

begin
  if ParamCount < 3 then begin
    ExitCode := 1;
    Write(ErrOutput, INTERNAL_NAME + ' v' + VERSION + ' by ' + AUTHOR + #13#10 + #13#10 +
    'Usage:' + #13#10 +
    '    SendSMS email password smstext [smswhere] [proxy:port]');
    exit;
  end;

  CoInitialize(nil);

  if not SendSMSviaGC(ParamStr(1), ParamStr(2), ParamStr(3), '', ifthen(ParamCount >= 4, ParamStr(4), ''), ifthen(ParamCount >= 5, ParamStr(5), ''), Answer) then begin
    Write(ErrOutput, Answer);
    ExitCode := 1;
  end
  else begin
    Write(Output, Answer);
    ExitCode := 0;
  end;

  CoUninitialize();
end.
(*
v0.1.1
[+] added DurationMinutes;
[~] changed DeltaMinutes form 2 to 4 minutes.

v0.1.2
[+] added xor-decoding of password;
[~!] fixed detection of successful creation of event;
[~] minor fixes.
*)
