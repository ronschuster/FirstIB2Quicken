object dlgAmazonOrder: TdlgAmazonOrder
  Left = 0
  Top = 0
  Caption = 'Amazon Order'
  ClientHeight = 432
  ClientWidth = 1075
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object Splitter1: TSplitter
    Left = 0
    Top = 210
    Width = 1075
    Height = 3
    Cursor = crVSplit
    Align = alTop
  end
  object RzDialogButtons1: TRzDialogButtons
    Left = 0
    Top = 396
    Width = 1075
    TabOrder = 0
  end
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 1075
    Height = 33
    Align = alTop
    TabOrder = 1
    object Label1: TLabel
      Left = 22
      Top = 8
      Width = 42
      Height = 13
      Caption = 'Order ID'
    end
    object edtOrderID: TEdit
      Left = 70
      Top = 5
      Width = 251
      Height = 21
      ReadOnly = True
      TabOrder = 0
      Text = 'edtOrderID'
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 33
    Width = 1075
    Height = 177
    Align = alTop
    Caption = 'Panel2'
    TabOrder = 2
    DesignSize = (
      1075
      177)
    object Label2: TLabel
      Left = 16
      Top = 8
      Width = 115
      Height = 13
      Caption = 'Credit card transactions'
    end
    object lvTransactions: TListView
      Left = 16
      Top = 24
      Width = 1042
      Height = 140
      Anchors = [akLeft, akTop, akRight, akBottom]
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
          Width = 525
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
        end>
      TabOrder = 0
      ViewStyle = vsReport
    end
  end
  object Panel3: TPanel
    Left = 0
    Top = 213
    Width = 1075
    Height = 183
    Align = alClient
    TabOrder = 3
    DesignSize = (
      1075
      183)
    object Label3: TLabel
      Left = 16
      Top = 8
      Width = 72
      Height = 13
      Caption = 'Amazon orders'
    end
    object lvOrders: TListView
      Left = 16
      Top = 24
      Width = 1036
      Height = 146
      Anchors = [akLeft, akTop, akRight, akBottom]
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
          Caption = 'Order ID'
          Width = 130
        end
        item
          Caption = 'Product Name'
          Width = 500
        end>
      TabOrder = 0
      ViewStyle = vsReport
    end
  end
end
