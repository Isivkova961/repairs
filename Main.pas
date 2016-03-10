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
  if not (Key in ['�'..'�','�'..'�','.',#32,#8]) then Key:=#0;
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
  TempleateFileName := ExtractFilePath(Application.ExeName) + '������.docx';
    try
      // ���� ��� �� ���������
      WordApp := CreateOleObject('Word.Application');

    except
      on E: Exception do
      begin
        ShowMessage('�� ������� ��������� Word!'#13#10 + E.Message);
        Exit;
      end;
    end;
  try
    Screen.Cursor := crHourGlass;

    // �������� ������ ��������� �� ������ �������
    Document := WordApp.Documents.Add(Template := TempleateFileName, NewTemplate := False);

    // �������� �������� �� ������
    ReplaceField(Document);
    // �� ��������� ���� Word ������, ������ ��� ������� � ������� �������
    WordApp.Visible := True;
    finally
    // ���������� ��� �������� ���������� Word.
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
      if CompareBm(BookmarkName, '�������') then
        Range.Text := eNumDoc.Text
      else
      if CompareBm(BookmarkName, '���') then
        Range.Text := eFIO.Text
      else
      if CompareBm(BookmarkName, '����') then
        Range.Text := deDate.Text
      else
      if CompareBm(BookmarkName, '�������') then
        Range.Text := eTelef.Text
      else
      if CompareBm(BookmarkName, '�����') then
        Range.Text := eSN.Text
      else
      if CompareBm(BookmarkName, '������������') then
        Range.Text := eName.Text
      else
      if CompareBm(BookmarkName, '�������������') then
        Range.Text := eNeispr.Text
      else
      if CompareBm(BookmarkName, '����������') then
        Range.Text := ePrim.Text
      else
      if CompareBm(BookmarkName, '�������') then
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
      if CompareBm(BookmarkName, '��������') then
        if cobGarant.Text='14 ����' then
          Range.Text := cobGarant.Text+' ������� � __ __ ____ �.'
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
      adotRem.FieldByName('���').AsString:=eFIO.Text;
      adotRem.FieldByName('���� ������').AsString:=deDate.Text;
      adotRem.FieldByName('����� ��������').AsString:=eNumDoc.Text;
      adotRem.FieldByName('�������').AsString:=eTelef.Text;
      adotRem.FieldByName('������������').AsString:=eName.Text;
      adotRem.FieldByName('�������� �����').AsString:=eSN.Text;
      adotRem.FieldByName('�������������').AsString:=eNeispr.Text;
      adotRem.FieldByName('������� �������').AsString:=s;
      adotRem.FieldByName('��������').AsString:=cobGarant.Text;
      adotRem.Post;
    end;
end;

end.
