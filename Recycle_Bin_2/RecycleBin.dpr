program RecycleBin;

uses
  Forms,
  NP in 'NP.pas' {MainFrm};

{$R *.res}

begin
  Application.Initialize;
  Application.Title := 'Recycle Bin';
  Application.CreateForm(TMainFrm, MainFrm);
  Application.Run;
end.
