object Form1: TForm1
  Left = 190
  Top = 234
  Width = 1000
  Height = 533
  Caption = 'Form1'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Splitter1: TSplitter
    Left = 641
    Top = 41
    Width = 5
    Height = 458
    Cursor = crHSplit
  end
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 992
    Height = 41
    Align = alTop
    TabOrder = 0
    object Label1: TLabel
      Left = 544
      Top = 14
      Width = 42
      Height = 13
      Caption = 'Category'
    end
    object btnReadChecking: TButton
      Left = 8
      Top = 8
      Width = 121
      Height = 25
      Caption = 'Read checking file'
      TabOrder = 0
      OnClick = btnReadCheckingClick
    end
    object btnCleanup: TButton
      Left = 128
      Top = 8
      Width = 75
      Height = 25
      Caption = 'Clean'#39'em up'
      TabOrder = 1
      OnClick = btnCleanupClick
    end
    object btnReadPayment: TButton
      Left = 200
      Top = 8
      Width = 129
      Height = 25
      Caption = 'Read bill payment  file'
      TabOrder = 2
      OnClick = btnReadPaymentClick
    end
    object btnWriteQIF: TButton
      Left = 456
      Top = 8
      Width = 75
      Height = 25
      Caption = 'Write QIF file'
      TabOrder = 3
      OnClick = btnWriteQIFClick
    end
    object btnReadCheckbook: TButton
      Left = 328
      Top = 8
      Width = 129
      Height = 25
      Caption = 'Read checkbook file'
      TabOrder = 4
      OnClick = btnReadCheckbookClick
    end
    object cboCategory: TwwDBComboBox
      Left = 592
      Top = 10
      Width = 209
      Height = 21
      ShowButton = True
      Style = csDropDown
      MapList = False
      AllowClearKey = False
      DropDownCount = 8
      HistoryList.Section = 'Categories'
      HistoryList.FileName = 'FirstIB2Quicken.INI'
      HistoryList.Enabled = True
      HistoryList.MRUEnabled = True
      HistoryList.MRUMaxSize = 5
      ItemHeight = 0
      Sorted = False
      TabOrder = 5
      UnboundDataType = wwDefault
    end
    object btnApply: TButton
      Left = 800
      Top = 8
      Width = 41
      Height = 25
      Caption = 'Apply'
      TabOrder = 6
      OnClick = btnApplyClick
    end
  end
  object pnlChecking: TPanel
    Left = 0
    Top = 41
    Width = 641
    Height = 458
    Align = alLeft
    BorderWidth = 8
    Caption = 'pnlChecking'
    TabOrder = 1
    OnDragOver = pnlCheckingDragOver
    object lvChecking: TListView
      Left = 9
      Top = 9
      Width = 623
      Height = 440
      Align = alClient
      Columns = <
        item
          Caption = 'Trans No'
          Width = 60
        end
        item
          Caption = 'Date'
          Width = 75
        end
        item
          Caption = 'Desc'
          Width = 75
        end
        item
          Caption = 'Memo'
          Width = 75
        end
        item
          Caption = 'Category'
        end
        item
          Caption = 'Debit'
          Width = 75
        end
        item
          Caption = 'Credit'
          Width = 75
        end
        item
          Caption = 'Balance'
          Width = 60
        end
        item
          Caption = 'Check No'
          Width = 60
        end
        item
          Caption = 'Fees'
        end>
      HideSelection = False
      MultiSelect = True
      RowSelect = True
      TabOrder = 0
      ViewStyle = vsReport
      OnColumnClick = lvCheckingColumnClick
      OnCompare = lvCheckingCompare
      OnDragDrop = lvCheckingDragDrop
      OnDragOver = lvCheckingDragOver
    end
  end
  object Panel3: TPanel
    Left = 646
    Top = 41
    Width = 346
    Height = 458
    Align = alClient
    BorderWidth = 8
    TabOrder = 2
    object lvPayments: TListView
      Left = 9
      Top = 9
      Width = 328
      Height = 440
      Align = alClient
      Columns = <
        item
          Caption = 'Date'
          Width = 75
        end
        item
          Caption = 'Amount'
          Width = 75
        end
        item
          Caption = 'Check No'
          Width = 64
        end
        item
          Caption = 'Payee'
          Width = 100
        end
        item
          Caption = 'Memo'
        end>
      DragMode = dmAutomatic
      RowSelect = True
      TabOrder = 0
      ViewStyle = vsReport
      OnColumnClick = lvPaymentsColumnClick
      OnCompare = lvPaymentsCompare
      OnKeyDown = lvPaymentsKeyDown
      OnStartDrag = lvPaymentsStartDrag
    end
  end
  object OpenDialog1: TOpenDialog
    Filter = '*.csv|*.csv|*.txt|*.txt'
    Left = 112
    Top = 48
  end
  object Timer1: TTimer
    Interval = 100
    OnTimer = Timer1Timer
    Left = 896
    Top = 8
  end
  object SaveDialog1: TSaveDialog
    DefaultExt = 'QIF'
    Filter = 'QIF files|*.qif'
    Options = [ofOverwritePrompt, ofHideReadOnly, ofPathMustExist, ofEnableSizing]
    Left = 152
    Top = 49
  end
end
