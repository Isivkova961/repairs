unit Rem;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, DBGrids, StdCtrls, Mask, ToolEdit, DBCtrls, GridsEh,
  DBGridEh, Menus,ComObj;

type
  TfRem = class(TForm)
    cebFIO: TCheckBox;
    cebName: TCheckBox;
    eFIO: TEdit;
    eName: TEdit;
    dbgRem: TDBGridEh;
    cebStatus: TCheckBox;
    cobStatus: TComboBox;
    lKol: TLabel;
    MainMenu1: TMainMenu;
    WORD1: TMenuItem;
    procedure FormShow(Sender: TObject);
    procedure NewFiltr;
    procedure FiltrBase;
    procedure Filtr(var str:string);
    procedure eFIOChange(Sender: TObject);
    procedure cebFIOClick(Sender: TObject);
    procedure cebNameClick(Sender: TObject);
    procedure eNameChange(Sender: TObject);
    procedure dbgRemKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure cebStatusClick(Sender: TObject);
    procedure cobStatusChange(Sender: TObject);
    procedure ReplaceField(const ADocument: OleVariant);
    procedure WORD1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fRem: TfRem;

implementation

uses RemDM;

{$R *.dfm}

procedure TfRem.FormShow(Sender: TObject);
begin
  NewFiltr;
  dmRem.adotRem.Filtered:=false;
  lKol.Caption:=IntToStr(dmRem.adotRem.RecordCount);
end;

procedure TfRem.NewFiltr;
begin
  cebFIO.Checked:=false;
  eFIO.Text:='';
  cebName.Checked:=false;
  eName.Text:='';
end;

procedure TfRem.FiltrBase;
var strFiltr:string;
begin
  Filtr(strFiltr);
  dmRem.adotRem.Filter:=strFiltr;
  dmRem.adotRem.Filtered:=(strFiltr<>'');
  lKol.Caption:=IntToStr(dmRem.adotRem.RecordCount);
end;

procedure TfRem.Filtr(var str:string);
var s:string;
begin
  if cebFIO.Checked then
    if s='' then
      s:=s+'ФИО LIKE ''*'+eFIO.Text+'*'''
    else
      s:=s+' AND ФИО LIKE ''*'+eFIO.Text+'*''';
  if cebName.Checked then
    if s='' then
      s:=s+'Наименование LIKE ''*'+eName.Text+'*'''
    else
      s:=s+' AND Наименование LIKE ''*'+eName.Text+'*''';
  if cebStatus.Checked then
    if s='' then
      s:=s+'Статус LIKE ''*'+cobStatus.Text+'*'''
    else
      s:=s+' AND Статус LIKE ''*'+cobStatus.Text+'*''';
  str:=s;
end;


procedure TfRem.eFIOChange(Sender: TObject);
begin
  if eFIO.Text<>'' then
    begin
      cebFIO.Checked:=true;
      FiltrBase;
    end;
end;

procedure TfRem.cebFIOClick(Sender: TObject);
begin
  if cebFIO.Checked=false then
    eFIO.Text:='';
  FiltrBase;
end;

procedure TfRem.cebNameClick(Sender: TObject);
begin
  if cebName.Checked=false then
    eName.Text:='';
  FiltrBase;
end;

procedure TfRem.eNameChange(Sender: TObject);
begin
  if eName.Text<>'' then
    begin
      cebName.Checked:=true;
      FiltrBase;
    end;
end;

procedure TfRem.dbgRemKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if key=13 then
    begin
      dmRem.adotRem.Edit;
      dmRem.adotRem.Post;
    end;
  if key=46 then
    if MessageDlg('Вы уверены что хотите удалить?',mtWarning,mbOkCancel,0)=mrOk then
      dmRem.adotRem.Delete;
  lKol.Caption:=IntToStr(dmRem.adotRem.RecordCount);
end;

procedure TfRem.cebStatusClick(Sender: TObject);
begin
  if cebStatus.Checked=false then
    cobStatus.Text:='';
  FiltrBase;
end;

procedure TfRem.cobStatusChange(Sender: TObject);
begin
  if cobStatus.Text<>'' then
    begin
      cebStatus.Checked:=true;
      FiltrBase;
    end;
end;



procedure TfRem.WORD1Click(Sender: TObject);
var TempleateFileName: string;
    WordApp, Document: OleVariant;
begin
  TempleateFileName := ExtractFilePath(Application.ExeName) + 'Шаблон.docx';
    try
      // Если нет то запускаем
      WordApp := CreateOleObject('Word.Application');

    except
      on E: Exception do
      begin
        ShowMessage('Не удалось запустить Word!'#13#10 + E.Message);
        Exit;
      end;
    end;
  try
    Screen.Cursor := crHourGlass;

    // Создание нового документа на основе шаблона
    Document := WordApp.Documents.Add(Template := TempleateFileName, NewTemplate := False);

    // Заменяем закладки на данные
    ReplaceField(Document);
    // По умолчание окно Word скрыто, делаем его видимым с готовым отчетом
    WordApp.Visible := True;
    finally
    // Необходимо для удаления экземпляра Word.
      WordApp := Unassigned;

    Screen.Cursor := crDefault;
  end;
end;

procedure TfRem.ReplaceField(const ADocument: OleVariant);
var i: Integer;
    BookmarkName: string;
    Range: OleVariant;
  function CompareBm(ABmName: string; const AName: string): Boolean;
  var j:integer;
  begin
    j := Pos('_', ABmName);
    if j > 0 then
      Delete(ABmName, j, Length(ABmName) - j + 1);
    Result := SameText(ABmName, AName);
  end;
begin
  for i := ADocument.Bookmarks.Count downto 1 do
    begin
      BookmarkName := ADocument.Bookmarks.Item(i).Name;
      Range := ADocument.Bookmarks.Item(i).Range;
      if CompareBm(BookmarkName, 'Договор') then
        Range.Text := dmRem.adotRem.FieldByName('Номер договора').AsString
      else
      if CompareBm(BookmarkName, 'ФИО') then
        Range.Text := dmRem.adotRem.FieldByName('ФИО').AsString
      else
      if CompareBm(BookmarkName, 'Дата') then
        Range.Text := dmRem.adotRem.FieldByName('Дата приема').AsString
      else
      if CompareBm(BookmarkName, 'Телефон') then
        Range.Text := dmRem.adotRem.FieldByName('Телефон').AsString
      else
      if CompareBm(BookmarkName, 'Номер') then
        Range.Text := dmRem.adotRem.FieldByName('Серийный номер').AsString
      else
      if CompareBm(BookmarkName, 'Наименование') then
        Range.Text := dmRem.adotRem.FieldByName('Наименование').AsString
      else
      if CompareBm(BookmarkName, 'Неисправность') then
        Range.Text := dmRem.adotRem.FieldByName('Неисправность').AsString
      else
      if CompareBm(BookmarkName, 'Примечание') then
        Range.Text := dmRem.adotRem.FieldByName('Примечание').AsString
      else
      if CompareBm(BookmarkName, 'Дефекты') then
        begin
          Range.Text := dmRem.adotRem.FieldByName('Внешние дефекты').AsString
        end
      else
      if CompareBm(BookmarkName, 'Гарантия') then
        if dmRem.adotRem.FieldByName('Гарантия').AsString='14 дней' then
          Range.Text := dmRem.adotRem.FieldByName('Гарантия').AsString+' начиная с __ __ ____ г.'
        else
          Range.Text := dmRem.adotRem.FieldByName('Гарантия').AsString;
    end;
end;

end.
