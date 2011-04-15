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
  VERSION       = '0.1.0';

var
  Answer: String;

function IfThen(Value: Boolean; const True: string; False: string = ''): string;
begin
  if Value then
    Result := True
  else
    Result := False;
end;

function HTTPRequest(Method: String; Url: String; Data: String; Referer: string; Header: TStringList; Proxy: string; var Error: Integer): String;
const
  SXH_PROXY_SET_PROXY = 2;
  HTTPREQUEST_SETCREDENTIALS_FOR_SERVER = 0;
  HTTPREQUEST_SETCREDENTIALS_FOR_PROXY = 1;

var
  ErrorCode, i: Integer;
  HttpReq: OLEVariant;
  Key, Value: string;
  ProxyServer, ProxyUsername, ProxyPassword: string;
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

  {ErrorCode := HttpRequest.setTimeouts(20000, 20000, 30000, 30000);
  if (ErrorCode <> S_OK) then begin
    Error := 3;
    Result := 'HTTP: Could not set timeouts.';
    exit;
  end;}

  if (Method = 'GET') and (Data <> '') then
    ErrorCode := HttpReq.Open(Method, Url + '?' + Data, false)
  else
    ErrorCode := HttpReq.Open(Method, Url, false);
  if (ErrorCode <> S_OK) then begin
    Error := 4;
    Result := 'HTTP: Could not send GET request.';
    exit;
  end;

  {if (Username <> '') or (Password <> '') then
  begin
    ErrorCode := HttpRequest.SetCredentials(
      Username, Password,
      HTTPREQUEST_SETCREDENTIALS_FOR_SERVER);
    if (ErrorCode <> S_OK) then begin
      Error := 5;
      Result := 'HTTP: Could not call SetCredentials().';
      exit;
    end;
  end;}

  HttpReq.SetRequestHeader('User-Agent', AUTHOR + '-' + INTERNAL_NAME + '-' + VERSION);
  if Referer <> '' then
    HttpReq.SetRequestHeader('Referer', Referer);

  // disable redirect
  HttpReq.Option(6) := false;

  if Method = 'GET' then
    ErrorCode := HttpReq.Send()
  else begin
    for i:= 0 to Header.Count - 1 do begin
      // Set HTTP Header
      Key := Header[i];
      Value := copy(Key, pos(': ', Key) + 2, MaxInt);
      Key := copy(Key, 1, pos(': ', Key) - 1);
      HttpReq.SetRequestHeader(Key, Value);
    end;
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

  if (HttpReq.Status >= 302) and (HttpReq.Status <= 304) and (not HttpReq.Option(6)) then begin
    Result := HTTPRequest(Method, HttpReq.GetResponseHeader('Location'), Data, Url, Header, Proxy, Error);
    exit;
  end;

  Result := HttpReq.ResponseText;
  if Trunc(HttpReq.Status / 100) = 2 then // 200, 201 ...
    Error := -1 // good response code
  else
    Error := HttpReq.Status;
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

function SendSMSviaGC(Email, Passwd, Title, Content, Where, Proxy: String; var Answer: String): boolean;
const
  ReminderMinutes = 1;
  DeltaMinutes = 2; // must by greater than ReminderMinutes
var
  url, data, Auth: string;
  Header: TStringList;
  err: Integer;
begin
  Result := false;

  url := 'https://www.google.com/accounts/ClientLogin';
	data := 'Email=' + email + '&Passwd=' + passwd + '&service=cl&source=' + AUTHOR + '-' + INTERNAL_NAME + '-' + VERSION;
  Header := TStringList.Create;
  Answer := HTTPRequest('POST', url, data, '', Header, Proxy, Err);
  if not Err = -1 then
    exit;

  Header.Text := Answer;
  Auth := Header.Values['Auth'];
  if Auth = '' then begin
    Answer := 'GC: ' + Header.Values['Error'];
    exit;
  end;

  url := 'http://www.google.com/calendar/feeds/default/private/full';
  data := '<atom:entry xmlns:atom="http://www.w3.org/2005/Atom">'+
  '  <atom:title type="text">' + T(Title) + '</atom:title>'+
  '  <atom:content type="text">' + T(Content) + '</atom:content>'+
  '  <gd:when xmlns:gd="http://schemas.google.com/g/2005" startTime="' + GetDateGC(IncMinutes(Now, DeltaMinutes)) + '" endTime="' + GetDateGC(IncMinutes(Now, DeltaMinutes)) + '">'+
  '    <gd:reminder minutes="' + inttostr(ReminderMinutes) + '" method="sms"/></gd:when>'+
  '  <gd:where xmlns:gd="http://schemas.google.com/g/2005" valueString="' + T(Where) + '"/></atom:entry>';
  Header.Clear;
  Header.NameValueSeparator := ':';
  Header.Add('authorization: GoogleLogin auth=' + Auth);
  Header.Add('Content-Type: application/atom+xml');
  Answer := HTTPRequest('POST', url, data, '', Header, Proxy, Err);

  Result := (Err = -1);
end;

begin
  if ParamCount < 3 then begin
    ExitCode := 1;
    Write(ErrOutput, INTERNAL_NAME + ' v' + VERSION + ' by ' + AUTHOR + #13#10 + #13#10 +
    'Usage:' + #13#10 +
    '    SendSMS email password smstext [smswhere proxy]');
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
