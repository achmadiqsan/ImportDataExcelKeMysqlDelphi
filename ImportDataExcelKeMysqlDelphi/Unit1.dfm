object Form1: TForm1
  Left = 224
  Top = 133
  Width = 928
  Height = 451
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poDesktopCenter
  OnActivate = FormActivate
  OnClose = FormClose
  PixelsPerInch = 96
  TextHeight = 13
  object ListView1: TListView
    Left = 16
    Top = 152
    Width = 433
    Height = 249
    Columns = <>
    TabOrder = 0
    ViewStyle = vsReport
  end
  object DBGrid1: TDBGrid
    Left = 464
    Top = 152
    Width = 433
    Height = 249
    DataSource = DataSource1
    TabOrder = 1
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'Tahoma'
    TitleFont.Style = []
  end
  object Edit1: TEdit
    Left = 24
    Top = 96
    Width = 121
    Height = 21
    TabOrder = 2
  end
  object Button1: TButton
    Left = 176
    Top = 96
    Width = 75
    Height = 25
    Caption = 'Browse'
    TabOrder = 3
    OnClick = Button1Click
  end
  object Button2: TButton
    Left = 264
    Top = 96
    Width = 89
    Height = 25
    Caption = 'Hapus ListView'
    TabOrder = 4
    OnClick = Button2Click
  end
  object Button3: TButton
    Left = 456
    Top = 96
    Width = 113
    Height = 25
    Caption = 'Simpan Database'
    TabOrder = 5
    OnClick = Button3Click
  end
  object Button4: TButton
    Left = 592
    Top = 96
    Width = 121
    Height = 25
    Caption = 'Kosongkan Database'
    TabOrder = 6
    OnClick = Button4Click
  end
  object Button5: TButton
    Left = 808
    Top = 96
    Width = 75
    Height = 25
    Caption = 'Close'
    TabOrder = 7
    OnClick = Button5Click
  end
  object ADOConnection1: TADOConnection
    Connected = True
    ConnectionString = 
      'Provider=MSDASQL.1;Persist Security Info=False;Data Source=impor' +
      'tdata'
    LoginPrompt = False
    Left = 128
    Top = 224
  end
  object ADOQuery1: TADOQuery
    Active = True
    Connection = ADOConnection1
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM ref_coa')
    Left = 232
    Top = 224
  end
  object DataSource1: TDataSource
    DataSet = ADOQuery1
    Left = 336
    Top = 248
  end
  object OpenDialog1: TOpenDialog
    Left = 192
    Top = 304
  end
end
