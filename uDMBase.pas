unit uDMBase;

interface

uses
  System.SysUtils, System.Classes, FireDAC.Stan.Intf, FireDAC.Stan.Option,
  FireDAC.Stan.Error, FireDAC.UI.Intf, FireDAC.Phys.Intf, FireDAC.Stan.Def,
  FireDAC.Stan.Pool, FireDAC.Stan.Async, FireDAC.Phys, FireDAC.Phys.SQLite,
  FireDAC.Phys.SQLiteDef, FireDAC.Stan.ExprFuncs, FireDAC.VCLUI.Wait,
  FireDAC.Stan.Param, FireDAC.DatS, FireDAC.DApt.Intf, FireDAC.DApt, Data.DB,
  FireDAC.Comp.DataSet, FireDAC.Comp.Client;

type
  TDMbase = class(TDataModule)
    MainConnection: TFDConnection;
    Exams: TFDQuery;
    procedure DataModuleCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  DMbase: TDMbase;

implementation

{%CLASSGROUP 'Vcl.Controls.TControl'}

{$R *.dfm}

var
  Database: string; // 'Data\advantage.sqlite3';

procedure TDMbase.DataModuleCreate(Sender: TObject);
begin
  MainConnection.DriverName := 'SQLite';
  MainConnection.Params.Values['Database'] := Database + '\advantage.sqlite3';
  MainConnection.Connected;
//  Assert(not MainConnection.Connected);
end;

initialization
  ForceDirectories(ExtractFilePath(ParamStr(0)) + 'Data\');
  Database := ExtractFilePath(ParamStr(0)) + 'Data\';
end.
