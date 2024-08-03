program Banks2Quicken;

uses
  Forms,
  frm_Main in 'frm_Main.pas' {frmMain},
  dlg_AmazonOrder in 'dlg_AmazonOrder.pas' {dlgAmazonOrder};

{$R *.RES}

begin
  Application.Initialize;
  Application.CreateForm(TfrmMain, frmMain);
  Application.CreateForm(TdlgAmazonOrder, dlgAmazonOrder);
  Application.Run;
end.
