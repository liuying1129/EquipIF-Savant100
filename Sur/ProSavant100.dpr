program ProSavant100;

uses
  Forms,
  UfrmMain in 'UfrmMain.pas' {frmMain},
  UCommFunction in 'UCommFunction.pas',
  PerlRegEx in 'PerlRegEx.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TfrmMain, frmMain);
  Application.ShowMainForm:=false;
  Application.Run;
end.
