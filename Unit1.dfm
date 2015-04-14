object Form1: TForm1
  Left = 30
  Top = 178
  Caption = 'Form1'
  ClientHeight = 609
  ClientWidth = 1211
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
  TextHeight = 16
  object Splitter1: TSplitter
    Left = 789
    Top = 89
    Width = 6
    Height = 520
    Margins.Left = 4
    Margins.Top = 4
    Margins.Right = 4
    Margins.Bottom = 4
    ExplicitTop = 50
    ExplicitHeight = 559
  end
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 1211
    Height = 89
    Margins.Left = 4
    Margins.Top = 4
    Margins.Right = 4
    Margins.Bottom = 4
    Align = alTop
    TabOrder = 0
    object Label1: TLabel
      Left = 837
      Top = 17
      Width = 55
      Height = 16
      Margins.Left = 4
      Margins.Top = 4
      Margins.Right = 4
      Margins.Bottom = 4
      Caption = 'Category'
    end
    object btnReadChecking: TButton
      Left = 10
      Top = 10
      Width = 122
      Height = 31
      Margins.Left = 4
      Margins.Top = 4
      Margins.Right = 4
      Margins.Bottom = 4
      Caption = 'Read checking file'
      TabOrder = 0
      OnClick = btnReadCheckingClick
    end
    object btnCleanup: TButton
      Left = 132
      Top = 10
      Width = 82
      Height = 31
      Margins.Left = 4
      Margins.Top = 4
      Margins.Right = 4
      Margins.Bottom = 4
      Caption = 'Clean'#39'em up'
      TabOrder = 1
      OnClick = btnCleanupClick
    end
    object btnReadPayment: TButton
      Left = 336
      Top = 10
      Width = 137
      Height = 31
      Margins.Left = 4
      Margins.Top = 4
      Margins.Right = 4
      Margins.Bottom = 4
      Caption = 'Read bill payment  file'
      TabOrder = 2
      OnClick = btnReadPaymentClick
    end
    object btnWriteQIF: TButton
      Left = 738
      Top = 10
      Width = 85
      Height = 31
      Margins.Left = 4
      Margins.Top = 4
      Margins.Right = 4
      Margins.Bottom = 4
      Caption = 'Write QIF file'
      TabOrder = 3
      OnClick = btnWriteQIFClick
    end
    object btnReadCheckbook: TButton
      Left = 474
      Top = 10
      Width = 135
      Height = 31
      Margins.Left = 4
      Margins.Top = 4
      Margins.Right = 4
      Margins.Bottom = 4
      Caption = 'Read checkbook file'
      TabOrder = 4
      OnClick = btnReadCheckbookClick
    end
    object cboCategory: TwwDBComboBox
      Left = 896
      Top = 12
      Width = 257
      Height = 24
      Margins.Left = 4
      Margins.Top = 4
      Margins.Right = 4
      Margins.Bottom = 4
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
      Left = 1152
      Top = 10
      Width = 50
      Height = 31
      Margins.Left = 4
      Margins.Top = 4
      Margins.Right = 4
      Margins.Bottom = 4
      Caption = 'Apply'
      TabOrder = 6
      OnClick = btnApplyClick
    end
    object btnReadDiscoverFile: TButton
      Left = 215
      Top = 10
      Width = 120
      Height = 31
      Margins.Left = 4
      Margins.Top = 4
      Margins.Right = 4
      Margins.Bottom = 4
      Caption = 'Read Discover file'
      TabOrder = 7
      OnClick = btnReadDiscoverFileClick
    end
    object btnReadDollarBankFile: TButton
      Left = 612
      Top = 10
      Width = 123
      Height = 31
      Margins.Left = 4
      Margins.Top = 4
      Margins.Right = 4
      Margins.Bottom = 4
      Caption = 'Read DollarBankfile'
      TabOrder = 8
      OnClick = btnReadDollarBankFileClick
    end
    object btnReadFinanceorksFile: TButton
      Left = 617
      Top = 49
      Width = 163
      Height = 31
      Margins.Left = 4
      Margins.Top = 4
      Margins.Right = 4
      Margins.Bottom = 4
      Caption = 'Read FinanceWorks File'
      TabOrder = 9
      OnClick = btnReadFinanceorksFileClick
    end
  end
  object pnlChecking: TPanel
    Left = 0
    Top = 89
    Width = 789
    Height = 520
    Margins.Left = 4
    Margins.Top = 4
    Margins.Right = 4
    Margins.Bottom = 4
    Align = alLeft
    BorderWidth = 8
    Caption = 'pnlChecking'
    TabOrder = 1
    OnDragOver = pnlCheckingDragOver
    ExplicitTop = 50
    ExplicitHeight = 559
    object lvChecking: TListView
      Left = 9
      Top = 9
      Width = 771
      Height = 502
      Margins.Left = 4
      Margins.Top = 4
      Margins.Right = 4
      Margins.Bottom = 4
      Align = alClient
      Columns = <
        item
          Caption = 'Trans No'
          Width = 74
        end
        item
          Caption = 'Date'
          Width = 92
        end
        item
          Caption = 'Desc'
          Width = 92
        end
        item
          Caption = 'Memo'
          Width = 92
        end
        item
          Caption = 'Category'
          Width = 62
        end
        item
          Caption = 'Debit'
          Width = 92
        end
        item
          Caption = 'Credit'
          Width = 92
        end
        item
          Caption = 'Balance'
          Width = 74
        end
        item
          Caption = 'Check No'
          Width = 74
        end
        item
          Caption = 'Fees'
          Width = 62
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
      ExplicitHeight = 541
    end
  end
  object Panel3: TPanel
    Left = 795
    Top = 89
    Width = 416
    Height = 520
    Margins.Left = 4
    Margins.Top = 4
    Margins.Right = 4
    Margins.Bottom = 4
    Align = alClient
    BorderWidth = 8
    TabOrder = 2
    ExplicitTop = 50
    ExplicitHeight = 559
    object lvPayments: TListView
      Left = 9
      Top = 9
      Width = 398
      Height = 502
      Margins.Left = 4
      Margins.Top = 4
      Margins.Right = 4
      Margins.Bottom = 4
      Align = alClient
      Columns = <
        item
          Caption = 'Date'
          Width = 92
        end
        item
          Caption = 'Amount'
          Width = 92
        end
        item
          Caption = 'Check No'
          Width = 79
        end
        item
          Caption = 'Payee'
          Width = 123
        end
        item
          Caption = 'Memo'
          Width = 62
        end>
      DragMode = dmAutomatic
      RowSelect = True
      TabOrder = 0
      ViewStyle = vsReport
      OnColumnClick = lvPaymentsColumnClick
      OnCompare = lvPaymentsCompare
      OnKeyDown = lvPaymentsKeyDown
      OnStartDrag = lvPaymentsStartDrag
      ExplicitHeight = 541
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
    Top = 32
  end
  object SaveDialog1: TSaveDialog
    DefaultExt = 'QIF'
    Filter = 'QIF files|*.qif'
    Options = [ofOverwritePrompt, ofHideReadOnly, ofPathMustExist, ofEnableSizing]
    Left = 152
    Top = 49
  end
end
