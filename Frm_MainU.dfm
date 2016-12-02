object Form1: TForm1
  Left = 0
  Top = 0
  ActiveControl = Edt_FileDir
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsSingle
  Caption = 'Excel to Laravel Seeder'
  ClientHeight = 592
  ClientWidth = 896
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 488
    Top = 8
    Width = 34
    Height = 13
    Caption = 'Result:'
  end
  object Shape1: TShape
    Left = 440
    Top = 11
    Width = 1
    Height = 566
    Pen.Color = clGray
  end
  object MakeModel: TButton
    Left = 24
    Top = 144
    Width = 385
    Height = 41
    Caption = 'Read File'
    TabOrder = 0
    OnClick = MakeModelClick
  end
  object Memo1: TMemo
    Left = 488
    Top = 27
    Width = 385
    Height = 543
    ScrollBars = ssVertical
    TabOrder = 1
  end
  object Edt_ModelName: TLabeledEdit
    Left = 24
    Top = 489
    Width = 121
    Height = 21
    EditLabel.Width = 62
    EditLabel.Height = 13
    EditLabel.Caption = 'Model Name:'
    TabOrder = 2
  end
  object Edt_FileDir: TLabeledEdit
    Left = 24
    Top = 41
    Width = 385
    Height = 21
    EditLabel.Width = 50
    EditLabel.Height = 13
    EditLabel.Caption = 'Excell File:'
    TabOrder = 3
  end
  object StringGrid1: TStringGrid
    Left = 24
    Top = 208
    Width = 385
    Height = 241
    ColCount = 2
    DefaultColWidth = 180
    RowCount = 2
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goEditing, goAlwaysShowEditor]
    TabOrder = 4
    ColWidths = (
      180
      180)
    RowHeights = (
      24
      24)
  end
  object Edt_WorkSheetNo: TLabeledEdit
    Left = 24
    Top = 101
    Width = 121
    Height = 21
    EditLabel.Width = 88
    EditLabel.Height = 13
    EditLabel.Caption = 'WorkSheet Index:'
    NumbersOnly = True
    TabOrder = 5
  end
  object Button1: TButton
    Left = 24
    Top = 529
    Width = 385
    Height = 41
    Caption = 'Make Seeder'
    TabOrder = 6
    OnClick = Button1Click
  end
end
