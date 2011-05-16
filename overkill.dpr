program overkill;

uses
  Windows,
  WinSvc,
  Variants,
  SysUtils,
  Classes,
  ActiveDS_TLB,
  ActiveX;

const
  SrvcName = 'OverKill';
  SrvcWait = 60000;
  SrvcLog = 'error.log';

var
  SrvcTable: SERVICE_TABLE_ENTRYA;
  SrvcStatus: SERVICE_STATUS;
  SrvcHandle: SERVICE_STATUS_HANDLE;
  FThread: Cardinal;

function ADsGetObject(lpszPathName: WideString; const riid: TGUID;
  out ppObject: Pointer): HRESULT; stdcall; external 'activeds.dll';

procedure WriteLog(AMessage: string);
var
  FFileHandle: TextFile;
  FFileName: string;
begin
  FFileName := ExtractFilePath(ParamStr(0)) + SrvcLog;
  AssignFile(FFileHandle, FFileName);
  if not FileExists(FFileName) then
  ReWrite(FFileHandle) else Append(FFileHandle);
  WriteLn(FFileHandle, FormatDateTime('[dd.mm.yyyy HH:nn:ss] ', Now) + AMessage);
  CloseFile(FFileHandle);
end;

function EnumCount(NewEnum: IInterface): Cardinal;
var
  FEnum: IEnumVariant;
  FCount: ULONG;
  FVariant: OleVariant;
  FIndex: Cardinal;
  FResult: HRESULT;
begin
  FIndex := 0;
  try
    FResult := NewEnum.QueryInterface(IEnumVariant, FEnum);
  except
    FResult := E_FAIL;
  end;
  while SUCCEEDED(FResult) do
  begin
    if not VarIsNull(FVariant) then FVariant := Unassigned;
    if FEnum.Next(1, FVariant, FCount) <> S_OK then Break;
    Inc(FIndex);
  end;
  FEnum := nil;
  Result := FIndex;
end;

function SrvcThread(Parameter: Pointer): DWORD; stdcall;
var
  FRPC: IADsFileServiceOperations;
  FEnum: IEnumVariant;
  FCount: ULONG;
  FItem: IADsSession;
  FVariant: OleVariant;
  FResult: HRESULT;
begin
  CoInitializeEx(nil, COINIT_MULTITHREADED);
  while (SrvcStatus.dwCurrentState = SERVICE_RUNNING) do
  begin
    try
      ADsGetObject('WinNT://./lanmanserver', IADsFileServiceOperations, Pointer(FRPC));
      FResult := FRPC.Sessions._NewEnum.QueryInterface(IEnumVariant, FEnum);
    except on E: Exception do
      begin
        WriteLog('Connect: ' + E.Message);
        FResult := E_FAIL;
      end;
    end;
    if EnumCount(FRPC.Sessions._NewEnum) >= 10 then
    while SUCCEEDED(FResult) do
    begin
      if not VarIsNull(FVariant) then FVariant := Unassigned;
      try
        if FEnum.Next(1, FVariant, FCount) <> S_OK then Break;
        IDispatch(FVariant).QueryInterface(IADsSession, Pointer(FItem));
      except on E: Exception do
        begin
          WriteLog('Enum: ' + E.Message);
          FResult := E_FAIL;
        end;
      end;
      try
        FRPC.Sessions.Remove(FItem.Name);
      except on E: Exception do
        begin
          WriteLog('Remove: ' + E.Message);
        end;
      end;
      FItem := nil;
    end;
    FEnum := nil;
    FRPC := nil;
    Sleep(SrvcWait);
  end;
  CoUninitialize;
  Result := 0;
end;

procedure SrvcCtrl(OpCode: Cardinal); stdcall;
begin
  case OpCode of
    SERVICE_CONTROL_PAUSE:
    begin
      SuspendThread(FThread);
      SrvcStatus.dwCurrentState := SERVICE_PAUSED;
      SrvcStatus.dwWin32ExitCode := 0;
      SetServiceStatus(SrvcHandle, SrvcStatus);
    end;
    SERVICE_CONTROL_CONTINUE:
    begin
      ResumeThread(FThread);
      SrvcStatus.dwCurrentState := SERVICE_RUNNING;
      SrvcStatus.dwWin32ExitCode := 0;
      SetServiceStatus(SrvcHandle, SrvcStatus);
    end;      
    SERVICE_CONTROL_STOP:
    begin
      SrvcStatus.dwCurrentState := SERVICE_STOPPED;
      SrvcStatus.dwWin32ExitCode := 0;
      SetServiceStatus(SrvcHandle, SrvcStatus);
      Exit;
    end;
    SERVICE_CONTROL_SHUTDOWN:
    begin
      SrvcStatus.dwCurrentState := SERVICE_STOPPED;
      SrvcStatus.dwWin32ExitCode := 0;
      SetServiceStatus(SrvcHandle, SrvcStatus);
      Exit;
    end;
  end;
end;

procedure SrvcMain(ArgC: DWORD; var ArgV: array of PChar); stdcall;
var
  iThread: Cardinal;
begin
  SrvcStatus.dwServiceType := SERVICE_WIN32_OWN_PROCESS;
  SrvcStatus.dwCurrentState := SERVICE_START_PENDING;
  SrvcStatus.dwControlsAccepted := SERVICE_ACCEPT_STOP
  or SERVICE_ACCEPT_SHUTDOWN or SERVICE_ACCEPT_PAUSE_CONTINUE;
  SrvcStatus.dwWin32ExitCode := 0;
  SrvcStatus.dwServiceSpecificExitCode := 0;
  SrvcStatus.dwCheckPoint := 0;
  SrvcStatus.dwWaitHint := 0;

  SrvcHandle := RegisterServiceCtrlHandler(SrvcName, @SrvcCtrl);

  SrvcStatus.dwCurrentState := SERVICE_RUNNING;
  SetServiceStatus(SrvcHandle, SrvcStatus);

  FThread := CreateThread(nil, 0, @SrvcThread, nil, 0, iThread);
  WaitForSingleObject(FThread, INFINITE);
  CloseHandle(FThread);
end;

procedure Install;
var
  FSCManager: SC_HANDLE;
  FSCService: SC_HANDLE;
  FBinExe: PChar;
begin
  try
    FBinExe := PChar(ParamStr(0));
    FSCManager := OpenSCManager(nil, nil, SC_MANAGER_ALL_ACCESS);
    FSCService := CreateService(FSCManager, PChar(SrvcName), PChar(SrvcName), SERVICE_ALL_ACCESS, SERVICE_WIN32_OWN_PROCESS, SERVICE_AUTO_START, SERVICE_ERROR_NORMAL, FBinExe, nil, nil, nil, nil, nil);
    CloseServiceHandle(FSCService);
  except on E: Exception do
    Writeln('ERROR: ' + E.Message);
  end;
end;

procedure Remove;
var
  FSCManager: SC_HANDLE;
  FSCService: SC_HANDLE;
begin
  try
    FSCManager := OpenSCManager(nil, nil, SC_MANAGER_ALL_ACCESS);
    FSCService := OpenService(FSCManager, PChar(SrvcName), SERVICE_ALL_ACCESS);
    DeleteService(FSCService);
    CloseServiceHandle(FSCService);
  except on E: Exception do
    Writeln('ERROR: ' + E.Message);
  end;
end;

begin
  if UpperCase(ParamStr(1)) = '/INSTALL' then Install
  else if UpperCase(ParamStr(1)) = '/REMOVE' then Remove
  else
  begin
    SrvcTable.lpServiceName := SrvcName;
    SrvcTable.lpServiceProc := @SrvcMain;
    StartServiceCtrlDispatcher(SrvcTable);
  end;
end.
