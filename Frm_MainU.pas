unit Frm_MainU;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.Grids, Vcl.StdCtrls,
  Vcl.ExtCtrls, System.Win.ComObj;

type
  TForm1 = class(TForm)
    MakeModel: TButton;
    Memo1: TMemo;
    Edt_ModelName: TLabeledEdit;
    Label1: TLabel;
    Edt_FileDir: TLabeledEdit;
    StringGrid1: TStringGrid;
    Edt_WorkSheetNo: TLabeledEdit;
    Button1: TButton;
    Shape1: TShape;
    procedure MakeModelClick(Sender: TObject);
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.Button1Click(Sender: TObject);
var
  XLApp, Sheet: OleVariant;
  x, y, r, c, i: Integer;
  RangeMatrix: Variant;
begin
  // Create Excel-OLE Object
  XLApp := CreateOleObject('Excel.Application');
  try
    // Hide Excel
    XLApp.Visible := False;

    // Open the Workbook
    XLApp.Workbooks.Open(Edt_FileDir.Text);

    Sheet := XLApp.WorkSheets[StrToInt(Edt_WorkSheetNo.Text)];

    y := Sheet.Usedrange.EntireColumn.count;
    x := Sheet.Usedrange.EntireRow.count;

    // Assign the Variant associated with the WorkSheet to the Delphi Variant
    RangeMatrix := Sheet.Usedrange.Value;

    Memo1.Clear;
    for r := 1 to x do
    begin
      Memo1.Lines.Add(Edt_ModelName.Text + '::create([');

      for i := 0 to StringGrid1.RowCount - 1 do
        if Length(StringGrid1.Rows[i][1]) > 0 then
        begin
          Memo1.Lines.Add(QuotedStr(StringGrid1.Rows[i][1]) + ' => ' +
            QuotedStr(Sheet.Cells.item[r, i + 1].Value) + ',');
        end;

      Memo1.Lines.Add(' ]);');
      Memo1.Lines.Add('');;
    end;

    // Unassign the Delphi Variant Matrix
    RangeMatrix := Unassigned;

  finally
    // Quit Excel
    if not VarIsEmpty(XLApp) then
    begin
      XLApp.DisplayAlerts := False;
      XLApp.Quit;
      XLApp := Unassigned;
      Sheet := Unassigned;
    end;
  end;

end;

procedure TForm1.MakeModelClick(Sender: TObject);
var
  XLApp, Sheet: OleVariant;
  y, r: Integer;
  RangeMatrix: Variant;
begin
  // Create Excel-OLE Object
  XLApp := CreateOleObject('Excel.Application');
  try
    // Hide Excel
    XLApp.Visible := False;

    // Open the Workbook
    XLApp.Workbooks.Open(Edt_FileDir.Text);

    Sheet := XLApp.WorkSheets[StrToInt(Edt_WorkSheetNo.Text)];

    y := Sheet.Usedrange.EntireColumn.count;

    // Set Stringgrid's row &col dimensions.
    StringGrid1.RowCount := y;

    // Assign the Variant associated with the WorkSheet to the Delphi Variant
    RangeMatrix := Sheet.Usedrange.Value;

    for r := 1 to y do
      StringGrid1.Cells[0, (r - 1)] := RangeMatrix[1, r];

    // Unassign the Delphi Variant Matrix
    RangeMatrix := Unassigned;

  finally
    // Quit Excel
    if not VarIsEmpty(XLApp) then
    begin
      XLApp.DisplayAlerts := False;
      XLApp.Quit;
      XLApp := Unassigned;
      Sheet := Unassigned;
    end;
  end;
end;

end.
