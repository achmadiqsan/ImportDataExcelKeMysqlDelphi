unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Grids, DBGrids, ComCtrls, DB, ADODB, ComObj;

type
  TForm1 = class(TForm)
    ListView1: TListView;
    DBGrid1: TDBGrid;
    Edit1: TEdit;
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    Button4: TButton;
    Button5: TButton;
    ADOConnection1: TADOConnection;
    ADOQuery1: TADOQuery;
    DataSource1: TDataSource;
    OpenDialog1: TOpenDialog;
    procedure Button5Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
  public
    { Public declarations }
    procedure bacaexcel (SheetIndex: Integer);
    procedure koneksi;
  end;

var
  Form1: TForm1;
  i,x: Integer;
  lvitem : TListItem;
  v: Variant;

implementation

{$R *.dfm}

procedure TForm1.bacaexcel(SheetIndex: Integer);
var
  Xlapp1, Sheet : Variant;
  i, MaxRow, MaxCol, x, y : Integer;
  Temp, Lokasi : string;
  Baris : TListItem;
begin
  Lokasi := OpenDialog1.FileName;
  Xlapp1 := CreateOleObject('excel.application');
  Xlapp1.Workbooks.open(Lokasi);
  Sheet := Xlapp1.WorkSheets[SheetIndex];
  MaxRow := Sheet.Usedrange.EntireRow.count;
  MaxCol := Sheet.Usedrange.EntireColumn.count;
  ListView1.Columns.Clear;

  for i := 1 to MaxCol do
  begin
    ListView1.Columns.Add.Caption := Sheet.cells.item[1,i].value;
  end;

  for x:= 2 to MaxRow do
  begin
    Baris := ListView1.Items.Add;
    Baris.Caption := Sheet.Cells.item[x,1];
    for y:=2 to MaxCol do
    begin
      Baris.SubItems.Add('');
      Temp:= Sheet.cells.item[x,y];
      ListView1.Items[x-2].SubItems.Strings[y-2]:= Temp;
    end;
  end;
  Xlapp1.Workbooks.close;
end;

procedure TForm1.Button5Click(Sender: TObject);
begin
  close;
end;

procedure TForm1.koneksi;
begin
  koneksi;
end;

procedure TForm1.Button1Click(Sender: TObject);
var
  i : Integer;
begin
  if OpenDialog1.Execute then
    begin
      Edit1.Text := OpenDialog1.FileName;
      bacaexcel(1);
      for i := 0 to ListView1.Columns.Count-1 do
      begin
        ListView1.Columns[i].Width := -2;
      end;
    end;
end;

procedure TForm1.Button2Click(Sender: TObject);
begin
  Edit1.Clear;
  ListView1.Clear;
  ADOQuery1.Open;
end;

procedure TForm1.Button3Click(Sender: TObject);
var
  i,x : Integer;
  lvitem : TListItem;
begin
  if ListView1.Items.Count = 0 then
  begin
    ShowMessage('Data Masih Kosong....');
    Exit;
  end;
  if ListView1.Columns.Count<>ADOQuery1.FieldCount then
    begin
      ShowMessage('Jumlah kolom tidak sama...');
      Exit;
    end;
  try
    for i := 1 to ListView1.Items.Count do
      begin
        ListView1.ItemIndex:= i - 1;
        lvitem := ListView1.Selected;
        ADOQuery1.Append;
        if ADOQuery1.Fields[0].DataType = ftDateTime then ADOQuery1.Fields[0].Value := StrToDate(lvitem.Caption)
        else if ADOQuery1.Fields[0].DataType = ftInteger then ADOQuery1.Fields[0].Value := StrToInt(lvitem.Caption)
        else ADOQuery1.Fields[0].Value := lvitem.Caption;
        for x:=1 to ADOQuery1.FieldCount-1 do
        begin
          if ADOQuery1.Fields[x].DataType = ftDateTime then ADOQuery1.Fields[x].Value := StrToDate(lvitem.SubItems[x-1])
          else if ADOQuery1.Fields[x].DataType = ftInteger then ADOQuery1.Fields[x].Value := StrToInt(lvitem.SubItems[x-1])
          else ADOQuery1.Fields[x].Value := lvitem.SubItems[x-1];
        end;
        ADOQuery1.Post;
        ADOQuery1.Close;
        ADOQuery1.Open;
      end;
      ListView1.Clear;
      ShowMessage('Data berhasil dipindahkan...');
  except
    ShowMessage('data gagal dipindahkan...');
  end;
end;

procedure TForm1.Button4Click(Sender: TObject);
begin
  ADOQuery1.First;
  while not ADOQuery1.Eof do
  begin
    ADOQuery1.Delete;
    ADOQuery1.Next;
    ADOQuery1.Close;
    ADOQuery1.Open;
  end;
end;

procedure TForm1.FormActivate(Sender: TObject);
begin
  ADOQuery1.Close;
  ADOQuery1.Open;
  Edit1.Clear;
  ListView1.Clear;
end;

procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Action:= caFree;
end;

end.
