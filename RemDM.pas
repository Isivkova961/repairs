unit RemDM;

interface

uses
  SysUtils, Classes, DB, ADODB;

type
  TdmRem = class(TDataModule)
    dsRem: TDataSource;
    adocRem: TADOConnection;
    adotRem: TADOTable;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  dmRem: TdmRem;

implementation

{$R *.dfm}

end.
