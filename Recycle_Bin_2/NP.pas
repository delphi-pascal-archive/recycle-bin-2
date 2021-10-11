unit NP;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms,
  Dialogs, Registry, XpMan, ShellApi, AppEvnts, StdCtrls;

type
  TMainFrm = class(TForm)
    ClearRecycleBin: TButton;
    ApplicationEvents: TApplicationEvents;
    OpenDialog: TOpenDialog;
    tx6: TLabel;
    tx8: TLabel;
    Frame: TGroupBox;
    tx1: TLabel;
    tx2: TLabel;
    tx3: TLabel;
    NameRecycleContextItem: TEdit;
    CommandRecycleContextItem: TEdit;
    tx4: TLabel;
    ChooseExe: TButton;
    AddToRecycleContext: TButton;
    DeleteBoxFromRecycleContext: TComboBox;
    UpdateList: TButton;
    DelFromRecycleContext: TButton;
    tx5: TLabel;
    AddToComputer: TCheckBox;
    tx7: TLabel;

    procedure ClearRecycleBinClick(Sender: TObject);
    procedure AddToComputerClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure DelFromRecycleContextClick(Sender: TObject);
    procedure UpdateListClick(Sender: TObject);
    procedure DeleteBoxFromRecycleContextCloseUp(Sender: TObject);
    procedure NameRecycleContextItemChange(Sender: TObject);
    procedure ApplicationEventsIdle(Sender: TObject; var Done: Boolean);
    procedure FormCreate(Sender: TObject);
    procedure tx8Click(Sender: TObject);
    procedure Edit1Change(Sender: TObject);
    procedure ChooseExeClick(Sender: TObject);
    procedure AddToRecycleContextClick(Sender: TObject);

  private

     R : TRegistry;

  public

  end;

var
  MainFrm: TMainFrm;

implementation

{$R *.dfm}

type
PSHQueryRBInfo = ^TSHQueryRBInfo;
TSHQueryRBInfo = packed record
cbSize: DWORD;
i64Size: Int64;
i64NumItems: Int64;
end;
const
shell32 = 'shell32.dll';

function SHQueryRecycleBin(szRootPath: PChar; SHQueryRBInfo: PSHQueryRBInfo): HResult;
stdcall; external shell32 Name 'SHQueryRecycleBinA';

function GetDllVersion(FileName: string): Integer;
var
InfoSize, Wnd: DWORD;
VerBuf: Pointer;
FI: PVSFixedFileInfo;
VerSize: DWORD;
begin
Result   := 0;
InfoSize := GetFileVersionInfoSize(PChar(FileName), Wnd);
if InfoSize <> 0 then
begin
GetMem(VerBuf, InfoSize);
try
if GetFileVersionInfo(PChar(FileName), Wnd, InfoSize, VerBuf) then
if VerQueryValue(VerBuf, '\', Pointer(FI), VerSize) then
Result := FI.dwFileVersionMS;
finally
FreeMem(VerBuf);
end;
end;
end;

procedure EmptyRecycleBin;
const
SHERB_NOCONFIRMATION = $00000001;
SHERB_NOPROGRESSUI = $00000002;
SHERB_NOSOUND = $00000004;
type
TSHEmptyRecycleBin = function(Wnd: HWND;
pszRootPath: PChar; dwFlags: DWORD): HRESULT;  stdcall;
var
SHEmptyRecycleBin: TSHEmptyRecycleBin;
LibHandle: THandle;
begin
LibHandle := LoadLibrary(PChar('Shell32.dll'));
if LibHandle <> 0 then @SHEmptyRecycleBin :=
GetProcAddress(LibHandle, 'SHEmptyRecycleBinA')
else
begin
MessageDlg('Failed to load Shell32.dll.', mtError, [mbOK], 0);
Exit;
end;
if @SHEmptyRecycleBin <> nil then
SHEmptyRecycleBin(Application.Handle, nil,
SHERB_NOCONFIRMATION or SHERB_NOPROGRESSUI or SHERB_NOSOUND);
FreeLibrary(LibHandle); @SHEmptyRecycleBin := nil;
end;

procedure TMainFrm.ClearRecycleBinClick(Sender: TObject);
begin
EmptyRecycleBin;
end;

procedure TMainFrm.AddToComputerClick(Sender: TObject);
begin
R := TRegistry.Create;
with R do begin
RootKey := HKEY_LOCAL_MACHINE;
if AddToComputer.Checked then
OpenKey('\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{645FF040-5081-101B-9F08-00AA002F954E}', True) else
DeleteKey('\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{645FF040-5081-101B-9F08-00AA002F954E}');
CloseKey;
Free;
end;
end;

procedure TMainFrm.FormShow(Sender: TObject);
begin
R:=TRegistry.Create;
R.RootKey := HKEY_LOCAL_MACHINE;
if R.KeyExists
('\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{645FF040-5081-101B-9F08-00AA002F954E}') then
AddToComputer.Checked := True else
AddToComputer.Checked := False;
R.CloseKey;
R.Free;
end;

procedure TMainFrm.DelFromRecycleContextClick(Sender: TObject);
var
i: Integer;
begin
if DeleteBoxFromRecycleContext.ItemIndex = -1 then
Exit else
begin
i := DeleteBoxFromRecycleContext.ItemIndex;
if DeleteBoxFromRecycleContext.ItemIndex  = -1 then
Exit else
R:=TRegistry.Create;
R.RootKey:=HKEY_CLASSES_ROOT;
r.DeleteKey('\\CLSID\{645FF040-5081-101B-9F08-00AA002F954E}\Shell\'+'\'+DeleteBoxFromRecycleContext.Items[DeleteBoxFromRecycleContext.itemindex]);
R.OpenKey('\\CLSID\{645FF040-5081-101B-9F08-00AA002F954E}\Shell\', False);
R.GetKeyNames(DeleteBoxFromRecycleContext.Items);
DelFromRecycleContext.Enabled := False;
DeleteBoxFromRecycleContext.Sorted := True;
R.CloseKey;
R.Free;
end;
end;

procedure TMainFrm.UpdateListClick(Sender: TObject);
begin
R:=TRegistry.Create;
R.RootKey:=HKEY_CLASSES_ROOT;
R.OpenKey('\CLSID\{645FF040-5081-101B-9F08-00AA002F954E}\Shell\', False);
R.GetKeyNames(DeleteBoxFromRecycleContext.Items);
DeleteBoxFromRecycleContext.Sorted:=true;
R.CloseKey;
R.Free;
end;

procedure TMainFrm.DeleteBoxFromRecycleContextCloseUp(Sender: TObject);
begin
if DeleteBoxFromRecycleContext.ItemIndex  = -1 then
DelFromRecycleContext.Enabled := False else
DelFromRecycleContext.Enabled := True;
end;

procedure TMainFrm.NameRecycleContextItemChange(Sender: TObject);
begin
if Length(CommandRecycleContextItem.Text) > 0 then
if Length(NameRecycleContextItem.Text) = 0 then
AddToRecycleContext.Enabled := False else
AddToRecycleContext.Enabled := True;
end;

procedure TMainFrm.ApplicationEventsIdle(Sender: TObject;
var Done: Boolean);
var
DllVersion: integer;
SHQueryRBInfo: TSHQueryRBInfo;
r: HResult;
begin
DllVersion := GetDllVersion(PChar(shell32));
if DllVersion >= $00040048 then
begin
FillChar(SHQueryRBInfo, SizeOf(TSHQueryRBInfo), #0);
SHQueryRBInfo.cbSize := SizeOf(TSHQueryRBInfo);
R := SHQueryRecycleBin(nil, @SHQueryRBInfo);
if r = s_OK then
begin
tx6.Caption := Format('Размер: %d' + ' Kb,' + ' Элементов: %d',
[SHQueryRBInfo.i64Size, SHQueryRBInfo.i64NumItems]);
end;
end;
end;

procedure TMainFrm.FormCreate(Sender: TObject);
begin
UpdateList.OnClick(Self);
end;

procedure TMainFrm.tx8Click(Sender: TObject);
begin
ShellExecute(Handle, nil, 'http://www.viacoding.mylivepage.ru/', nil,nil, Sw_ShowNormal);
end;

procedure TMainFrm.Edit1Change(Sender: TObject);
begin
if Length(NameRecycleContextItem.Text) > 0 then
if Length(CommandRecycleContextItem.Text) = 0 then
AddToRecycleContext.Enabled := False else
AddToRecycleContext.Enabled := True;
end;

procedure TMainFrm.ChooseExeClick(Sender: TObject);
begin
if OpenDialog.Execute then
begin
CommandRecycleContextItem.Text := OpenDialog.FileName;
end;
end;

procedure TMainFrm.AddToRecycleContextClick(Sender: TObject);
begin
try
R:=TRegistry.Create;
R.RootKey:=HKEY_CLASSES_ROOT;
if R.OpenKey('\CLSID\{645FF040-5081-101B-9F08-00AA002F954E}\Shell\' + NameRecycleContextItem.Text + '\Command',true) then
begin
R.WriteString('',CommandRecycleContextItem.Text);
end;
DelFromRecycleContext.Enabled := False;
NameRecycleContextItem.Clear;
CommandRecycleContextItem.Clear;
except
end;
end;

end.
