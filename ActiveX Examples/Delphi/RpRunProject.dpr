program RpRunProject;

uses
  Forms,
  RpFormUnit in 'RpFormUnit.pas' {frmReport};

{$R *.RES}

begin
  Application.Initialize;
  Application.CreateForm(TfrmReport, frmReport);
  Application.Run;
end.
