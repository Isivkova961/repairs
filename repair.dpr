program repair;

uses
  Forms,
  Main in 'Main.pas' {fRepair},
  RemDM in 'RemDM.pas' {dmRem: TDataModule},
  Rem in 'Rem.pas' {fRem};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TfRepair, fRepair);
  Application.CreateForm(TdmRem, dmRem);
  Application.CreateForm(TfRem, fRem);
  Application.Run;
end.
