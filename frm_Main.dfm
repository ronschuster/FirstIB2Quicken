object frmMain: TfrmMain
  Left = 30
  Top = 178
  Caption = 'Banks2Quicken'
  ClientHeight = 503
  ClientWidth = 1104
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  Menu = MainMenu1
  OldCreateOrder = False
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Splitter1: TSplitter
    Left = 689
    Top = 36
    Width = 5
    Height = 467
  end
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 1104
    Height = 36
    Align = alTop
    TabOrder = 0
    object lblFilenameChk: TLabel
      Left = 88
      Top = 4
      Width = 3
      Height = 13
      Caption = ' '
    end
    object lblFilenamePymt: TLabel
      Left = 88
      Top = 18
      Width = 3
      Height = 13
      Caption = ' '
    end
    object btnCleanup: TButton
      Left = 9
      Top = 5
      Width = 67
      Height = 25
      Caption = 'Clean'#39'em up'
      TabOrder = 0
      OnClick = btnCleanupClick
    end
    object Panel2: TPanel
      Left = 786
      Top = 1
      Width = 317
      Height = 34
      Align = alRight
      BevelOuter = bvNone
      TabOrder = 1
      object Label1: TLabel
        Left = 8
        Top = 10
        Width = 42
        Height = 13
        Caption = 'Category'
      end
      object cboCategory: TwwDBComboBox
        Left = 56
        Top = 6
        Width = 209
        Height = 21
        ShowButton = True
        Style = csDropDown
        MapList = False
        AllowClearKey = False
        DropDownCount = 20
        HistoryList.Section = 'Categories'
        HistoryList.FileName = 'FirstIB2Quicken.INI'
        HistoryList.Enabled = True
        HistoryList.MRUEnabled = True
        HistoryList.MRUMaxSize = 10
        ItemHeight = 0
        Sorted = False
        TabOrder = 0
        UnboundDataType = wwDefault
      end
      object btnApply: TButton
        Left = 269
        Top = 4
        Width = 41
        Height = 25
        Caption = 'Apply'
        TabOrder = 1
        OnClick = btnApplyClick
      end
    end
  end
  object pnlChecking: TPanel
    Left = 0
    Top = 36
    Width = 689
    Height = 467
    Align = alLeft
    BorderWidth = 8
    Caption = 'pnlChecking'
    TabOrder = 1
    OnDragOver = pnlCheckingDragOver
    object lvChecking: TListView
      Left = 9
      Top = 9
      Width = 671
      Height = 449
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
      OnCustomDrawItem = lvCheckingCustomDrawItem
      OnDblClick = lvCheckingDblClick
      OnDragDrop = lvCheckingDragDrop
      OnDragOver = lvCheckingDragOver
    end
  end
  object Panel3: TPanel
    Left = 694
    Top = 36
    Width = 410
    Height = 467
    Align = alClient
    BorderWidth = 8
    TabOrder = 2
    object lvPayments: TListView
      Left = 9
      Top = 9
      Width = 392
      Height = 449
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
      OnChange = lvPaymentsChange
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
    Top = 32
  end
  object SaveDialog1: TSaveDialog
    DefaultExt = 'QIF'
    Filter = 'QIF files|*.qif'
    Options = [ofOverwritePrompt, ofHideReadOnly, ofPathMustExist, ofEnableSizing]
    Left = 152
    Top = 49
  end
  object MainMenu1: TMainMenu
    Left = 112
    Top = 136
    object File1: TMenuItem
      Caption = 'File'
      object Read1: TMenuItem
        Caption = 'Read'
        object mniChecking: TMenuItem
          Caption = 'Checking'
          OnClick = btnReadCheckingClick
        end
        object mniDiscover: TMenuItem
          Caption = 'Discover'
          OnClick = btnReadDiscoverFileClick
        end
        object mniBillpayment: TMenuItem
          Caption = 'Bill payment'
          OnClick = btnReadPaymentClick
        end
        object mniCheckbook: TMenuItem
          Caption = 'Checkbook'
          OnClick = btnReadCheckbookClick
        end
        object mniFinanceWorks: TMenuItem
          Caption = 'FinanceWorks'
          OnClick = btnReadFinanceorksFileClick
        end
        object mniChaseVisa: TMenuItem
          Caption = 'Chase Visa'
          OnClick = mniChaseVisaClick
        end
        object mniCitiVisa: TMenuItem
          Caption = 'citi Visa'
          OnClick = mniCitiVisaClick
        end
        object mniMint: TMenuItem
          Caption = 'Mint'
          OnClick = mniMintClick
        end
        object mniAmazonOrders: TMenuItem
          Caption = 'Amazon Orders'
          OnClick = mniAmazonOrdersClick
        end
        object mniAmazonReturns: TMenuItem
          Caption = 'Amazon Returns'
          OnClick = mniAmazonReturnsClick
        end
      end
      object mniWriteQIF: TMenuItem
        Caption = 'Write QIF'
        OnClick = btnWriteQIFClick
      end
      object mniConvertMintcategoriestoQuicken: TMenuItem
        Caption = 'Convert Mint categories to Quicken'
        OnClick = mniConvertMintcategoriestoQuickenClick
      end
      object mniProcessAmazonTransactions: TMenuItem
        Caption = 'Process Amazon Transactions'
        OnClick = mniProcessAmazonTransactionsClick
      end
      object mniExit: TMenuItem
        Caption = 'Exit'
        OnClick = mniExitClick
      end
    end
  end
end
