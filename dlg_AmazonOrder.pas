unit dlg_AmazonOrder;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ComCtrls,
  Vcl.ExtCtrls, RzPanel, RzDlgBtn;

type
  TdlgAmazonOrder = class(TForm)
    RzDialogButtons1: TRzDialogButtons;
    Panel1: TPanel;
    Label1: TLabel;
    edtOrderID: TEdit;
    Panel2: TPanel;
    Label2: TLabel;
    lvTransactions: TListView;
    Panel3: TPanel;
    Label3: TLabel;
    lvOrders: TListView;
    Splitter1: TSplitter;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  dlgAmazonOrder: TdlgAmazonOrder;

implementation

{$R *.dfm}

end.
