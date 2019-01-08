 program QTI;

uses
  Vcl.Forms,
  MainForm in 'MainForm.pas' {frmMain},
  uDMBase in 'uDMBase.pas' {DMbase: TDataModule},
  XMLStructure in 'XMLStructure.pas';

{$R *.res}

begin
  Application.Initialize;
  ReportMemoryLeaksOnShutdown :=  True;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TfrmMain, frmMain);
  Application.CreateForm(TDMbase, DMbase);
  Application.Run;
end.


