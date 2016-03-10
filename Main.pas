unit Main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, Mask, ToolEdit,ComObj, CheckLst, Menus;

type
  TfRepair = class(TForm)
    lFIO: TLabel;
    lDate: TLabel;
    lName: TLabel;
    lSN: TLabel;
    lNeispr: TLabel;
    lPrim: TLabel;
    lDefekt: TLabel;
    lGarant: TLabel;
    eFIO: TEdit;
    eName: TEdit;
    eSN: TEdit;
    eNeispr: TEdit;
    ePrim: TEdit;
    cobGarant: TComboBox;
    bbForm: TBitBtn;
    bbExit: TBitBtn;
    deDate: TDateEdit;
    lTelef: TLabel;
    eTelef: TEdit;
    lNumDoc: TLabel;
    eNumDoc: TEdit;
    clbDefekt: TCheckListBox;
    mmMain: TMainMenu;
    N1: TMenuItem;
    procedure bbExitClick(Sender: TObject);
    procedure eFIOKeyPress(Sender: TObject; var Key: Char);
    procedure FormShow(Sender: TObject);
    procedure bbFormClick(Sender: TObject);
    procedure ReplaceField(const ADocument: OleVariant);
    procedure eNumDocKeyPress(Sender: TObject; var Key: Char);
    procedure N1Click(Sender: TObject);
    procedure ZapisBase;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fRepair: TfRepair;
  garant:string;

implementation

uses RemDM, Rem;

{$R *.dfm}

procedure TfRepair.bbExitClick(Sender: TObject);
begin
  Close;
end;

procedure TfRepair.eFIOKeyPress(Sender: TObject; var Key: Char);
begin
  if not (Key in ['а'..'я','А'..'Я','.',#32,#8]) then Key:=#0;
end;

procedure TfRepair.FormShow(Sender: TObject);
var F:TextFile;
    s:string;
begin
  AssignFile(F,'count.dat');
  if FileExists('count.dat') then
    Reset(F)
  else
    Rewrite(F);
  Read(F,s);
  eNumDoc.Text:=s;
  CloseFile(F);
  deDate.Date:=Date;
end;

procedure TfRepair.bbFormClick(Sender: TObject);
var TempleateFileName: string;
    WordApp, Document: OleVariant;
    F:TextFile;
    num:integer;
begin
  ZapisBase;
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
  AssignFile(F,'count.dat');
  Rewrite(F);
  num:=StrToInt(eNumDoc.Text)+1;
  Write(F,IntToStr(num));
  CloseFile(F);
  Close;
end;


procedure TfRepair.ReplaceField(const ADocument: OleVariant);
var i,k: Integer;
    BookmarkName,s: string;
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
        Range.Text := eNumDoc.Text
      else
      if CompareBm(BookmarkName, 'ФИО') then
        Range.Text := eFIO.Text
      else
      if CompareBm(BookmarkName, 'Дата') then
        Range.Text := deDate.Text
      else
      if CompareBm(BookmarkName, 'Телефон') then
        Range.Text := eTelef.Text
      else
      if CompareBm(BookmarkName, 'Номер') then
        Range.Text := eSN.Text
      else
      if CompareBm(BookmarkName, 'Наименование') then
        Range.Text := eName.Text
      else
      if CompareBm(BookmarkName, 'Неисправность') then
        Range.Text := eNeispr.Text
      else
      if CompareBm(BookmarkName, 'Примечание') then
        Range.Text := ePrim.Text
      else
      if CompareBm(BookmarkName, 'Дефекты') then
        begin
          s:='';
          for k:=0 to clbDefekt.Count-1 do
            if clbDefekt.Checked[k] then
              if s='' then
                s:=s+clbDefekt.Items[k]
              else
                s:=s+','+clbDefekt.Items[k];
          Range.Text := s;
        end
      else
      if CompareBm(BookmarkName, 'Гарантия') then
        if cobGarant.Text='14 дней' then
          Range.Text := cobGarant.Text+' начиная с __ __ ____ г.'
        else
          Range.Text := cobGarant.Text;
    end;
end;

procedure TfRepair.eNumDocKeyPress(Sender: TObject; var Key: Char);
begin
  if not (Key in ['0'..'9',#8]) then Key:=#0;
end;

procedure TfRepair.N1Click(Sender: TObject);
begin
  fRem.ShowModal;
end;

procedure TfRepair.ZapisBase;
var s:string;
    k:integer;
begin
  for k:=0 to clbDefekt.Count-1 do
    if clbDefekt.Checked[k] then
      if s='' then
        s:=s+clbDefekt.Items[k]
      else
        s:=s+','+clbDefekt.Items[k];
  with dmRem do
    begin
      adotRem.Insert;
      adotRem.FieldByName('ФИО').AsString:=eFIO.Text;
      adotRem.FieldByName('Дата приема').AsString:=deDate.Text;
      adotRem.FieldByName('Номер договора').AsString:=eNumDoc.Text;
      adotRem.FieldByName('Телефон').AsString:=eTelef.Text;
      adotRem.FieldByName('Наименование').AsString:=eName.Text;
      adotRem.FieldByName('Серийный номер').AsString:=eSN.Text;
      adotRem.FieldByName('Неисправность').AsString:=eNeispr.Text;
      adotRem.FieldByName('Внешние дефекты').AsString:=s;
      adotRem.FieldByName('Гарантия').AsString:=cobGarant.Text;
      adotRem.Post;
    end;
end;

end.
