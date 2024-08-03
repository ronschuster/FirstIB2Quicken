unit frm_Main;
{
01/24/10   RJS   Refactored btnReadPaymentClick: pulled QIF reading code out
                 into new procedure ReadQIFFile.
05/13/11   RJS   Added support for Category.
05/13/11   RJS   Added Read Discover button
12/15/12   RJS   Added Read DollarBankfile button
03/30/14   RJS   For DollarBank, strip leading "ONL ##### ' where ##### is a 5-digit number.
                 Changed btnReadCheckingClick to search from the first line for Column headings line rather than having a hard-coded line number.
                 Convert to Delphi XE.
04/13/15   RJS   Added Read FinanceWorks File
04/04/17   RJS   Added coloring of items in checking list matching  item selected in payment list.
06/27/20   RJS   Replaced individual function buttons with menu. Added ReadCSVFile to combine common code.
06/27/20   RJS   Changed ReadCSVFile calls to use anonymous methods. Added Chase Visa and citi Visa formats.
03/22/21   RJS   deleted DollarBank
                 added label for filename
                 updated file format for ChaseVisa
                 turned off EurekaLog. it prevented see exceptions I wanted to see.
04/09/22   RJS   Removed Chase Checking
                 Added support for Mint
08/03/24   RJS   Added support for processing Amazon orders
}


interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ComCtrls, ExtCtrls, Mask, wwdbedit, Wwdotdot, Wwdbcomb, Menus;

type
  TChkTrans = class
    TransactionNumber: string;
    Date: TDateTime;
    Description: string;
    Memo: string;
    AmountDebit: Double;
    AmountCredit: Double;
    Balance: Double;
    CheckNumber: Integer;
    Fees: Double;
    Category: string;
    OrderID: string;
    AmznTransProcessed: Boolean;
    function ValueByName(Name: string): Variant;
  end;

(*
    TPmtTrans = class
      Date: TDateTime;
      Amt: Double;
      Payee: string;
      CheckNumber: Integer;
    end;
*)

  TCopyTransToListItem = procedure (T:TChkTrans; L:TListItem) of Object;
  TProcessLineProc = reference to procedure (S: string; Trans: TChkTrans);

  TfrmMain = class(TForm)
    OpenDialog1: TOpenDialog;
    Panel1: TPanel;
    pnlChecking: TPanel;
    lvChecking: TListView;
    btnCleanup: TButton;
    Splitter1: TSplitter;
    Panel3: TPanel;
    lvPayments: TListView;
    Timer1: TTimer;
    SaveDialog1: TSaveDialog;
    MainMenu1: TMainMenu;
    File1: TMenuItem;
    Read1: TMenuItem;
    mniWriteQIF: TMenuItem;
    mniExit: TMenuItem;
    mniChecking: TMenuItem;
    mniDiscover: TMenuItem;
    mniBillpayment: TMenuItem;
    mniCheckbook: TMenuItem;
    mniFinanceWorks: TMenuItem;
    Panel2: TPanel;
    Label1: TLabel;
    cboCategory: TwwDBComboBox;
    btnApply: TButton;
    mniChaseVisa: TMenuItem;
    mniCitiVisa: TMenuItem;
    lblFilenameChk: TLabel;
    mniMint: TMenuItem;
    mniConvertMintcategoriestoQuicken: TMenuItem;
    mniAmazonOrders: TMenuItem;
    mniAmazonReturns: TMenuItem;
    lblFilenamePymt: TLabel;
    mniProcessAmazonTransactions: TMenuItem;
    procedure btnReadCheckingClick(Sender: TObject);
    procedure lvCheckingCompare(Sender: TObject; Item1, Item2: TListItem;
      Data: Integer; var Compare: Integer);
    procedure btnCleanupClick(Sender: TObject);
    procedure lvCheckingColumnClick(Sender: TObject; Column: TListColumn);
    procedure btnReadPaymentClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure lvPaymentsColumnClick(Sender: TObject; Column: TListColumn);
    procedure lvPaymentsCompare(Sender: TObject; Item1, Item2: TListItem;
      Data: Integer; var Compare: Integer);
    procedure lvPaymentsKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure lvCheckingDragOver(Sender, Source: TObject; X, Y: Integer;
      State: TDragState; var Accept: Boolean);
    procedure lvPaymentsStartDrag(Sender: TObject;
      var DragObject: TDragObject);
    procedure lvCheckingDragDrop(Sender, Source: TObject; X, Y: Integer);
    procedure Timer1Timer(Sender: TObject);
    procedure pnlCheckingDragOver(Sender, Source: TObject; X, Y: Integer;
      State: TDragState; var Accept: Boolean);
    procedure btnWriteQIFClick(Sender: TObject);
    procedure btnReadCheckbookClick(Sender: TObject);
    procedure btnApplyClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure btnReadDiscoverFileClick(Sender: TObject);
    procedure btnReadFinanceorksFileClick(Sender: TObject);
    procedure lvCheckingCustomDrawItem(Sender: TCustomListView; Item: TListItem;
      State: TCustomDrawState; var DefaultDraw: Boolean);
    procedure lvPaymentsChange(Sender: TObject; Item: TListItem;
      Change: TItemChange);
    procedure mniExitClick(Sender: TObject);
    procedure mniChaseCheckingClick(Sender: TObject);
    procedure mniChaseVisaClick(Sender: TObject);
    procedure mniCitiVisaClick(Sender: TObject);
    procedure mniMintClick(Sender: TObject);
    procedure mniConvertMintcategoriestoQuickenClick(Sender: TObject);
    procedure mniAmazonOrdersClick(Sender: TObject);
    procedure mniAmazonReturnsClick(Sender: TObject);
    procedure mniProcessAmazonTransactionsClick(Sender: TObject);
    procedure lvCheckingDblClick(Sender: TObject);
  private
    { Private declarations }
    slFileText: TStringList;
    Delimiter: Char;
    LoadingFile: Boolean;  // while loading file in the right-hand pane, suppress custom draw on left-hand pane
    slFields: TStringList;
    procedure CleanUpItem (Item: TListItem);
    procedure CopyChkTransToListItem(T:TChkTrans; L:TListItem);
    procedure CopyPmtTransToListItem(T:TChkTrans; L:TListItem);
    procedure ReadQIFFile(Filename: string; ListView: TListView;
      CopyTransToListItem: TCopyTransToListItem);
    function ClearListView(LV: TListView): Boolean;
    procedure ParseLine(S: string);
    function MatchingTransaction(CT, Pmt: TChkTrans): Boolean;
    procedure ReadCSVFile(HeaderLine: string; AProcessLine: TProcessLineProc);
    procedure CopyAmznOrderToListItem(T: TChkTrans; L: TListItem);
    function GetField(Index: Integer): string; overload;
    procedure SetDefaultPaymentsCaptions;
    procedure CheckFileFormat(Expected, Received: string);
    function GetAmazonOrdersLoaded: Boolean;
  public
    { Public declarations }
    property Fields[Index: Integer]: string read GetField;
    property AmazonOrdersLoaded: Boolean read GetAmazonOrdersLoaded;
  end;

var
  frmMain: TfrmMain;

implementation

{$R *.DFM}

uses
  StStrL, RegularExpressions, Generics.Collections, System.UITypes{, System.JSON,
  System.IOUtils}, System.StrUtils, CodeSiteLogging, dlg_AmazonOrder,
  System.Math;

var
  ChkSortField, PmtSortField: string;
  ChkSortAsc, PmtSortAsc: Boolean;
  DragPmt: TChkTrans;
  DragIndex: Integer;


procedure TfrmMain.CleanUpItem (Item: TListItem);
{for any of the specific Desc, the Memo is cleaned up
 and put into Desc and the Memo is cleared. For any other Desc
 Desc and Memo are left alone.}
var
  Desc, Memo: string;
//  NewMemo: string;
  I: Integer;
  Trans: TChkTrans;
(*
  procedure SkipDigits;
  begin
    while (Pos <= Length(Memo)) and (Memo[Pos] in ['0'..'9']) do
      Inc(Pos);
  end;

  procedure NewMemoToDesc;
  begin
    Trans.Description := NewMemo + Copy(Memo, Pos, 255);
    Trans.Memo := '';
  end;

  procedure NewMemoToMemo;
  begin
    Trans.Memo := NewMemo + Copy(Memo, Pos, 255);
  end;

  procedure CopyUntil(C: Char);
  begin
    while (Pos <= Length(Memo)) and (Memo[Pos] <> C) do begin
      NewMemo := NewMemo + Memo[Pos];
      Inc(Pos);
    end;
  end;

  procedure SkipUntil(C: Char);
  begin
    while (Pos <= Length(Memo)) and (Memo[Pos] <> C) do
      Inc(Pos);
  end;

  procedure StripTrailing(N: Integer);
  begin
    Delete(Memo, Length(Memo)-(N-1), N);
  end;
*)
const
  PhraseSet1: array[1..7] of string = (
    //FirstIB
    'External Deposit',
    'External Withdrawal',
    'Point Of Sale Deposit',
    'Point Of Sale Withdrawal',
    //DollarBank
    'POS-PIN',
    'POS-MC',
    'DIR');

  PhraseSet2: array[1..2] of string = (
    'ATM Withdrawal',
    'ATM Deposit');
begin
  Trans := TChkTrans(Item.Data);
  Desc := Trans.Description;
  Memo := Trans.Memo;

{ PhraseSet1: For all of these phrases, if phrase found at beginning of Desc, delete the phrase
  from Desc. If remaining Desc is not blank, add a space. Append Memo to end of Desc }
  for I := 1 to High(PhraseSet1) do
    if Pos(PhraseSet1[I], Desc) = 1 then begin
      Delete(Desc, 1, Length(PhraseSet1[I])+1);
      if Desc <> '' then
        Desc := Desc + ' ';
      Trans.Description := Desc + Memo;
      Trans.Memo := '';
      Break;
    end;

{ PhraseSet2: For all of these phrases, if phrase found at beginning of Desc,
  leave the phrase at beginning of Desc, and remove the remainder of the Desc
  and insert it and the beginning of Memo }
  for I := 1 to 2 do
    if Pos(PhraseSet2[I], Desc) = 1 then begin
      Trans.Description := PhraseSet2[I];
      Trans.Memo := Copy(Desc, Length(PhraseSet2[I])+2, 255) + ' ' + Memo;
      Break;
    end;

// for DollarBank, strip leading "ONL ##### ' where ##### is a 5-digit number
  if TRegEx.IsMatch(Desc, '^ONL \d\d\d\d\d ') then
    Trans.Description := Copy(Desc, 11, 255);

(*
  NewMemo := '';
  Pos := 1;
  if Desc = 'Purchase' then begin
    if Memo[1] = '#' then begin
      Inc(Pos, 2);
      SkipDigits;
      Inc(Pos, 1);
    end;
    StripTrailing(8);
    NewMemoToDesc;
  end
  else if Desc = 'ACH Deposit' then begin
    CopyUntil('(');
    Inc(Pos, 1);
    SkipDigits;
    Inc(Pos, 1);
    NewMemoToMemo;
  end
  else if Desc = 'ACH Pre-Authorized W/D' then begin
    CopyUntil('(');
    Inc(Pos, 1);
    SkipDigits;
    Inc(Pos, 1);
    NewMemoToDesc;
  end
  else if Desc = 'POS Withdrawal' then begin
    Inc(Pos, 2);
    SkipUntil('#');
    Inc(Pos, 1);
    SkipDigits;
    Inc(Pos, 1);
    StripTrailing(15);
    NewMemoToDesc;
  end
  else if (Desc = 'Withdrawal Via ATM')
       or (Desc = 'Deposit Via ATM') then begin
    Inc(Pos, 2);
    SkipDigits;
    Inc(Pos, 1);
    StripTrailing(15);
    NewMemoToMemo;
  end
  else if (Desc = 'Withdrawal Checking') and (Memo[1] <> '#') then
    NewMemoToDesc;
*)
  CopyChkTransToListItem(Trans, Item);
end;

function AmtToStr(A: Double): string;
begin
  if A = 0 then
    Result := ''
  else
    Result := Format('%f',[A]);
end;

procedure TfrmMain.CopyChkTransToListItem(T:TChkTrans; L:TListItem);
begin
  with T do begin
    L.Caption := TransactionNumber;
    with L.SubItems do begin
      Clear;
      Add(DateToStr(Date));
      Add(Description);
      Add(Memo);
      Add(Category);
      Add(AmtToStr(AmountDebit));
      Add(AmtToStr(AmountCredit));
      Add(AmtToStr(Balance));
      Add(IntToStr(CheckNumber));
      Add(AmtToStr(Fees));
    end;
  end;
end;

function StrToAmt(S: string): Double;
begin
  if Trim(S) = '' then
    Result := 0
  else
    Result := StrToFloat(FilterL(S,','));
end;

procedure TfrmMain.CopyPmtTransToListItem(T:TChkTrans; L:TListItem);
begin
  with T do begin
    L.Caption := DateToStr(Date);
    with L.SubItems do begin
      Clear;
      Add(AmtToStr(AmountDebit));
      Add(IntToStr(CheckNumber));
      Add(Description);
      Add(Memo);
    end;
  end;
end;

procedure TfrmMain.CopyAmznOrderToListItem(T:TChkTrans; L:TListItem);
begin
  with T do begin
    L.Caption := DateToStr(Date);
    with L.SubItems do begin
      Clear;
      Add(AmtToStr(AmountDebit));
      Add(OrderID);
      Add(Description);
    end;
  end;
end;

function TfrmMain.ClearListView(LV: TListView): Boolean;
var
  I: Integer;
begin
  if LV.Items.Count = 0 then
    Result := True
  else begin
    Result := MessageDlg('This will clear all transations currently in the list. Do you want to continue?', mtConfirmation, [mbYes, mbNo], 0) = mrYes;
    if Result then begin
      for I := 0 to LV.Items.Count-1 do
        TChkTrans(LV.Items[I]).Free;
      LV.Items.Clear;
    end;
  end;
end;

procedure TfrmMain.btnReadCheckingClick(Sender: TObject);

  procedure ProcessLine(S: string);
  var
    ListItem: TListItem;
    Trans: TChkTrans;
  begin
    if Length(S) = 0 then
      Exit;

    ParseLine(S);
    Trans := TChkTrans.Create;
    with Trans do begin
      TransactionNumber := Fields[0];
      Date := StrToDate(Fields[1]);
      Description := Fields[2];
      Memo := Fields[3];
      AmountDebit := StrToAmt(Fields[4]);
      AmountCredit := StrToAmt(Fields[5]);
      Balance := StrToAmt(Fields[6]);
      CheckNumber := StrToIntDef(Fields[7],0);
      Fees := StrToAmt(Fields[8]);

      with lvChecking do begin
        ListItem := Items.Add;
        ListItem.Data := Trans;
        CopyChkTransToListItem(Trans, ListItem);
      end;
    end;
  end;

var
  I: Integer;
  Filename: string;
  ColumnHeadingsFound: Boolean;
begin
  OpenDialog1.Filter := '*.csv|*.csv|*.qif|*.qif';
  if OpenDialog1.Execute then begin
    Filename := OpenDialog1.Filename;
    lblFilenameChk.Caption := Filename;
    if SameText(ExtractFileEXt(Filename), '.qif') then
      ReadQIFFile(Filename, lvChecking, CopyChkTransToListItem)
    else begin
      if not ClearListView(lvChecking) then
        Exit;
      slFileText.LoadFromFile(Filename);
      Delimiter := ',';
      ColumnHeadingsFound := False;
      for I := 0 to slFileText.Count-1 do
        if not ColumnHeadingsFound then begin
          if FilterL(slFileText[I],' ') = 'TransactionNumber,Date,Description,Memo,AmountDebit,AmountCredit,Balance,CheckNumber,Fees' then
            ColumnHeadingsFound := True;
        end
        else
          ProcessLine(slFileText[I]);
      if not ColumnHeadingsFound then
        raise Exception.Create('Invalid file input format. Column headings line not found.');
    end
  end;
end;

procedure TfrmMain.btnReadDiscoverFileClick(Sender: TObject);
begin
  ReadCSVFile('Trans. Date,Post Date,Description,Amount,Category',
    procedure (S: string; Trans: TChkTrans)
    begin
      ParseLine(S);
      with Trans do begin
        Date := StrToDate(Fields[0]);
        Description := Fields[2];
        AmountDebit := -1 * StrToAmt(Fields[3]);
        if AmountDebit > 0 then begin
          AmountCredit := AmountDebit;
          AmountDebit := 0;
        end;
        Category := Fields[4];
      end;
    end
  );
end;

procedure TfrmMain.lvCheckingCompare(Sender: TObject; Item1, Item2: TListItem;
  Data: Integer; var Compare: Integer);
var
  Trans1, Trans2: TChkTrans;
  Field1, Field2: Variant;
begin
  Trans1 := TChkTrans(Item1.Data);
  Trans2 := TChkTrans(Item2.Data);
  Field1 := Trans1.ValueByName(ChkSortField);
  Field2 := Trans2.ValueByName(ChkSortField);
  if Field1= Field2 then
    Compare := 0
  else if Field1 < Field2 then
    Compare := -1
  else
    Compare := 1;
  if not ChkSortAsc then
    Compare := -1 * Compare;
end;

procedure TfrmMain.lvCheckingCustomDrawItem(Sender: TCustomListView;
  Item: TListItem; State: TCustomDrawState; var DefaultDraw: Boolean);
const
  clMatch = $CCFFCC;  // color of matching list items
var
  Accept: Boolean;
begin
  if lvPayments.Selected = nil then
    Accept := False
  else
    Accept := MatchingTransaction(TChkTrans(Item.Data), TChkTrans(lvPayments.Selected.Data));
  if Accept then
    lvChecking.Canvas.Brush.Color := clMatch
  else
    lvChecking.Canvas.Brush.Color := clWindow;
end;

procedure TfrmMain.btnCleanupClick(Sender: TObject);
var
  I: Integer;
begin
  for I := 0 to lvChecking.Items.Count-1 do
    CleanUpItem(lvChecking.Items[I]);
end;

procedure TfrmMain.lvCheckingDblClick(Sender: TObject);
var
  sOrderID: string;
  ListItem: TListItem;
  TransList, OrderList: TList<TChkTrans>;
  I: Integer;
begin
  TransList := nil;
  OrderList := nil;
  try
    TransList := TList<TChkTrans>.Create;
    OrderList := TList<TChkTrans>.Create;
    if lvChecking.Selected <> nil then begin
      sOrderID :=  TChkTrans(lvChecking.Selected.Data).OrderID;
      if sOrderID <> '' then begin
        dlgAmazonOrder.edtOrderID.Text := sOrderID;
        dlgAmazonOrder.lvTransactions.Clear;
        dlgAmazonOrder.lvOrders.Clear;
        for I := 0 to lvChecking.Items.Count - 1 do
          if TChkTrans(lvChecking.Items[I].Data).OrderID = sOrderID then
            TransList.Add(TChkTrans(lvChecking.Items[I].Data));
        for I := 0 to lvPayments.Items.Count - 1 do
          with TChkTrans(lvPayments.Items[I].Data) do
            if (OrderID = sOrderID) and (AmountDebit <> 0) then
              OrderList.Add(TChkTrans(lvPayments.Items[I].Data));
        for I := 0 to TransList.Count - 1 do begin
          with dlgAmazonOrder.lvTransactions do begin
            ListItem := Items.Add;
            CopyChkTransToListItem(TransList[I], ListItem);
          end;
        end;
        for I := 0 to OrderList.Count - 1 do
          with dlgAmazonOrder.lvOrders do begin
            ListItem := Items.Add;
            CopyAmznOrderToListItem(OrderList[I], ListItem);
          end;
        dlgAmazonOrder.ShowModal;
      end;
    end;
  finally
    TransList.Free;
    OrderList.Free;
  end;
end;

procedure TfrmMain.lvCheckingColumnClick(Sender: TObject;
  Column: TListColumn);
begin
  lvChecking.SortType := stNone;
  if Column.Caption = ChkSortField then
    ChkSortAsc := not ChkSortAsc
  else
    ChkSortAsc := True;
  ChkSortField := Column.Caption;
  lvChecking.SortType := stData;
end;

procedure TfrmMain.ReadQIFFile(Filename: string; ListView: TListView; CopyTransToListItem: TCopyTransToListItem);
var
  I, J, NbrTrans: Integer;
  Trans: TChkTrans;
  ListItem: TListItem;
  S: string;
begin
  if not ClearListView(ListView) then
    Exit;
  if ListView = lvChecking then
    lblFilenameChk.Caption := Filename
  else
    lblFilenamePymt.Caption := Filename;
  slFileText.LoadFromFile(Filename);
  CheckFileFormat('!Type:Bank', slFileText[0]);

  NbrTrans := 0;
  for I := 1 to slFileText.Count-1 do
    if slFileText[I] = '^' then
      Inc(NbrTrans);

  J := 1;
  for I := 0 to NbrTrans-1 do begin
    Trans := TChkTrans.Create;
    with Trans do begin
      while slFileText[J] <> '^' do begin
        S := slFileText[J];
        case S[1] of
          'D': Date := StrToDate(Copy(S, 2, 10));
          'T': begin
                 AmountDebit := StrToAmt(Copy(S, 2, 20));
                 if AmountDebit > 0 then begin
                   AmountCredit := AmountDebit;
                   AmountDebit := 0;
                 end;
               end;
          'P': Description := Copy(S, 2, 80);
          'L': Category := Copy(S, 2, 80);
          'M': Memo := Copy(S, 2, 80);
          'N': if CharInSet(S[2], ['0'..'9']) then
                 CheckNumber := StrToIntDef(Copy(S, 2, 10), 0);
        end;
        Inc(J);
      end;
      Inc(J);
      with ListView do begin
        ListItem := Items.Add;
        ListItem.Data := Trans;
        CopyTransToListItem(Trans, ListItem);
      end;
    end;
  end;
end;

procedure TfrmMain.ParseLine(S: string);

  function GetField(InStr: string; var Pos: Integer): string;
  //get the next field from a comma or tab-delimited string
  var
    C: Char;
    State: Integer;
  begin
    State := 0;
    Result := '';
    while Pos <= Length(InStr) do begin
      C := InStr[Pos];
      case State of
        0: if C = Delimiter then begin
             Inc(Pos);
             Exit;
           end
           else if C = '"' then
             State := 1
           else
             Result := Result + C;
        1: case C of
             '"': State := 0;
             else Result := Result + C;
           end;
      end;
      Inc(Pos);
    end;
  end;

begin
  var Pos := 1;
  slFields.Clear;
  while Pos <= Length(S) do
    slFields.Add(GetField(S, Pos));
end;

procedure TfrmMain.SetDefaultPaymentsCaptions;
begin
  with lvPayments do begin
    Columns[2].Caption := 'Check No';
    Columns[3].Caption := 'Payee';
  end;
end;

procedure TfrmMain.btnReadPaymentClick(Sender: TObject);
// 12/31/2015,"Grace C & MA Church",$562.00,JX1F0-1RGX0,DDA - 2206,Complete,Charitable Donations

  procedure ProcessLine(S: string);
  var
    ListItem: TListItem;
    Trans: TChkTrans;
    Status: string;
  begin
    if Length(S) = 0 then
      Exit;
    ParseLine(S);
    Trans := TChkTrans.Create;
    with Trans do begin
      Date := StrToDate(Fields[0]);
      Description := Fields[1];
      AmountDebit := -StrToAmt(FilterL(Fields[2],'$'));
      Status := Fields[5];
      Category := Fields[6];
      if Status = 'Complete' then
        with lvPayments do begin
          ListItem := Items.Add;
          ListItem.Data := Trans;
          CopyPmtTransToListItem(Trans, ListItem);
        end
      else
        Trans.Free;
    end;
  end;

var
  Filename: string;
  I: Integer;
begin
  OpenDialog1.Filter := '*.csv|*.csv';
  if OpenDialog1.Execute then begin
    SetDefaultPaymentsCaptions;
    Filename := OpenDialog1.Filename;
    if not ClearListView(lvPayments) then
      Exit;
    lblFilenamePymt.Caption := Filename;
    slFileText.LoadFromFile(Filename);
    Delimiter := ',';
    CheckFileFormat('Deliver by,Paid to,Amount,Confirmation No.,Paid from,Status,Category', slFileText[0]);
    try
      LoadingFile := True;
      for I := 1 to slFileText.Count-1 do
        ProcessLine(slFileText[I]);
    finally
      LoadingFile := False;
    end;
  end
end;

procedure TfrmMain.btnReadCheckbookClick(Sender: TObject);
  procedure ProcessLine(S: string);
  var
    ListItem: TListItem;
    Trans: TChkTrans;
  begin
    if Length(S) = 0 then
      Exit;
    ParseLine(S);
    Trans := TChkTrans.Create;
    with Trans do begin
      CheckNumber := StrToIntDef(Fields[0],0);
      Date := StrToDate(Fields[1]);
      Description := Fields[2];
      Memo := Fields[3];
      AmountDebit := -StrToAmt(Fields[4]);

      with lvPayments do begin
        ListItem := Items.Add;
        ListItem.Data := Trans;
        CopyPmtTransToListItem(Trans, ListItem);
      end;
    end;
  end;

var
  Filename: string;
  I: Integer;
begin
  OpenDialog1.Filter := 'Tab delimited text files (*.txt)|*.txt';
  if OpenDialog1.Execute then begin
    SetDefaultPaymentsCaptions;
    Filename := OpenDialog1.Filename;
    if SameText(ExtractFileEXt(Filename), '.qif') then
      ReadQIFFile(Filename, lvPayments, CopyPmtTransToListItem)
    else begin
      if not ClearListView(lvPayments) then
        Exit;
      lblFilenamePymt.Caption := Filename;
      slFileText.LoadFromFile(Filename);
      Delimiter := #9;
      CheckFileFormat('Chk#'#9'Date'#9'Payee'#9'Memo'#9'Amt', slFileText[0]);
      try
        LoadingFile := True;
        for I := 1 to slFileText.Count-1 do
          ProcessLine(slFileText[I]);
      finally
        LoadingFile := False;
      end;
    end
  end
end;

function TfrmMain.GetAmazonOrdersLoaded: Boolean;
begin
  Result := lvPayments.Columns[2].Caption = 'Order ID';
end;

function TfrmMain.GetField(Index: Integer): string;
begin
  if Index < slFields.Count then
    Result := slFields[Index]
  else
    Result := '';
end;

procedure TfrmMain.FormCreate(Sender: TObject);
begin
  slFileText := TStringList.Create;
  slFields := TStringList.Create;
end;

procedure TfrmMain.FormDestroy(Sender: TObject);
begin
  slFileText.Free;
  slFields.Free;
end;

procedure TfrmMain.lvPaymentsChange(Sender: TObject; Item: TListItem;
  Change: TItemChange);
begin
  if not LoadingFile then
    lvChecking.Repaint;
end;

procedure TfrmMain.lvPaymentsColumnClick(Sender: TObject;
  Column: TListColumn);
begin
  lvPayments.SortType := stNone;
  if Column.Caption = PmtSortField then
    PmtSortAsc := not PmtSortAsc
  else
    PmtSortAsc := True;
  PmtSortField := Column.Caption;
  lvPayments.SortType := stData;
end;

procedure TfrmMain.lvPaymentsCompare(Sender: TObject; Item1, Item2: TListItem;
  Data: Integer; var Compare: Integer);
var
  Trans1, Trans2: TChkTrans;
  Field1, Field2: Variant;
begin
  Trans1 := TChkTrans(Item1.Data);
  Trans2 := TChkTrans(Item2.Data);
  Field1 := Trans1.ValueByName(PmtSortField);
  Field2 := Trans2.ValueByName(PmtSortField);
  if Field1= Field2 then
    Compare := 0
  else if Field1 < Field2 then
    Compare := -1
  else
    Compare := 1;
  if not PmtSortAsc then
    Compare := -1 * Compare;
end;

procedure TfrmMain.lvPaymentsKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var
  L: TListItem;
  I: Integer;
  PT: TChkTrans;
begin
  if Key = VK_DELETE then begin
    L := lvPayments.Selected;
    if L = nil then
      Exit;
    I := lvPayments.Items.IndexOf(L);
    PT := TChkTrans(L.Data);
    if MessageDlg('Are you sure you want to delete '+
        L.Caption+' '+AmtToStr(PT.AmountDebit)+' '+PT.Description+'?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then begin
      PT.Free;
      lvPayments.Items.Delete(I);
    end;
  end;
end;

function TfrmMain.MatchingTransaction(CT, Pmt: TChkTrans): Boolean;
{ returns true if the checking transaction CT matches the payment Pmt for the
  purposes of highlighting valid drag targets. }
begin
  Result := (CT.AmountDebit = Pmt.AmountDebit)
             and (CT.Date >= Pmt.Date);
end;

procedure TfrmMain.CheckFileFormat(Expected, Received: string);
begin
  Received := Trim(Received);
  if Received <> Expected then
    raise Exception.Create('Invalid file input format.'#13#13'Expected: ' + Expected + #13#13'Received: ' + Received);
end;

procedure TfrmMain.ReadCSVFile(HeaderLine: string; AProcessLine: TProcessLineProc);

  procedure ProcessLine(S: string);
  var
    ListItem: TListItem;
    Trans: TChkTrans;
  begin
    if Length(S) = 0 then
      Exit;
    Trans := TChkTrans.Create;
    with Trans do begin
      AProcessLine(S, Trans);
      with lvChecking do begin
        ListItem := Items.Add;
        ListItem.Data := Trans;
        CopyChkTransToListItem(Trans, ListItem);
      end;
    end;
  end;

var
  I: Integer;
  Filename: string;
begin
  OpenDialog1.Filter := '*.csv|*.csv';
  if OpenDialog1.Execute then begin
    Filename := OpenDialog1.Filename;
    if not ClearListView(lvChecking) then
      Exit;
    lblFilenameChk.Caption := Filename;
    slFileText.LoadFromFile(Filename);
    Delimiter := ',';
    while Trim(slFileText[0]) = '' do
      slFileText.Delete(0);
    CheckFileFormat(HeaderLine, slFileText[0]);
    for I := 1 to slFileText.Count-1 do
      ProcessLine(Trim(slFileText[I]));
  end
end;

procedure TfrmMain.mniExitClick(Sender: TObject);
begin
  Close
end;

procedure TfrmMain.lvCheckingDragOver(Sender, Source: TObject; X, Y: Integer;
  State: TDragState; var Accept: Boolean);
var
  L: TListItem;
begin
  L := lvChecking.GetItemAt(X,Y);
  if (L = nil) or (DragPmt = nil) then
    Exit;
  Accept := MatchingTransaction(TChkTrans(L.Data), DragPmt);
end;

procedure TfrmMain.lvPaymentsStartDrag(Sender: TObject;
  var DragObject: TDragObject);
var
  L: TListItem;
begin
  L := lvPayments.Selected;
  if L = nil then
    DragPmt := nil
  else begin
    DragPmt := TChkTrans(L.Data);
    DragIndex := lvPayments.Items.IndexOf(L);
  end;
end;

procedure TfrmMain.lvCheckingDragDrop(Sender, Source: TObject; X,
  Y: Integer);
var
  CT: TChkTrans;
  L: TListItem;
begin
  L := lvChecking.GetItemAt(X,Y);
  if (L = nil) or (DragPmt = nil) then
    Exit;
  CT := TChkTrans(L.Data);
  CT.Description := DragPmt.Description;
  CT.Memo := DragPmt.Memo;
  CopyChkTransToListItem(CT, L);

  lvPayments.Items.Delete(DragIndex);
  DragPmt.Free;
end;

procedure TfrmMain.Timer1Timer(Sender: TObject);
var
  P: TPoint;
  OK: Boolean;
begin
  OK := False;
  if lvPayments.Dragging then begin
    P := pnlChecking.ScreenToClient(Mouse.CursorPos);
    if (P.X >= 0) and (P.X < pnlChecking.Width) then begin
      if (P.Y >= 0) and (P.Y <= pnlChecking.BorderWidth) then begin
        //Scroll up
        SendMessage(lvChecking.Handle, WM_VSCROLL, SB_LINEUP, 0);
        OK := True;  //keep timer running
      end
      else if (P.Y >= (pnlChecking.Height - pnlChecking.BorderWidth)) and (P.Y < pnlChecking.Height) then begin
        //Scroll down
        SendMessage(lvChecking.Handle, WM_VSCROLL, SB_LINEDOWN, 0);
        OK := True;  //keep timer running
      end;
    end;
  end;

  if not OK then begin
    Timer1.Enabled := False;
    Exit;
  end;
end;

procedure TfrmMain.pnlCheckingDragOver(Sender, Source: TObject; X,
  Y: Integer; State: TDragState; var Accept: Boolean);
begin
  Accept := False;
  Timer1.Enabled := True;
end;

procedure TfrmMain.btnWriteQIFClick(Sender: TObject);
var
  I: Integer;
  Trans: TChkTrans;
  sAmount: string;
begin
  slFileText.Clear;
  slFileText.Add('!Type:Bank');
  if SaveDialog1.Execute then begin
    for I := 0 to lvChecking.Items.Count-1 do begin
      Trans := TChkTrans(lvChecking.Items[I].Data);
      with Trans do begin
        slFileText.Add('D'+DateToStr(Date));
        sAmount := Format('%.2f', [AmountDebit + AmountCredit]);
        slFileText.Add('U'+sAmount);
        slFileText.Add('T'+sAmount);
        slFileText.Add('P'+Description);
        slFileText.Add('L'+Category);
        if Memo <> '' then
          slFileText.Add('M'+Memo);
        if CheckNumber <> 0 then
          slFileText.Add('N'+IntToStr(CheckNumber));
        slFileText.Add('^');
      end;
    end;
    slFileText.SaveToFile(SaveDialog1.Filename)
  end;
end;

procedure TfrmMain.btnReadFinanceorksFileClick(Sender: TObject);
begin
  ReadCSVFile('Date, Account Name, Check #, Transaction, Category, Note, Expense, Deposit',
    procedure (S: string; Trans: TChkTrans)
    begin
      ParseLine(S);
      with Trans do begin
        Date := StrToDate(Fields[0]);
        CheckNumber :=  StrToIntDef(Fields[2],0);
        Description := Fields[3];
        Category := Fields[4];
        Memo := Fields[5];
        AmountDebit := StrToAmt(Fields[6]);
        AmountCredit := StrToAmt(Fields[7]);
      end;
    end
  );
end;

function ISOStrToDateTime(const S: string): TDateTime;
{This function converts an ISO 8601 date-time string to a TDateTime. ISO 8601 is an
 international standard for the formatting of dates and times. It has been adopted
 by most of the major countries around the world, either directly or as equivalent
 national standards. This function currently supports the following formats:
   yyyy-mm-ddThh:mm:ss
   yyyy-mm-dd hh:mm:ss
   yyyy-mm-ddThh:mm
   yyyy-mm-dd hh:mm
   yyyy-mm-dd
   yyyy-mm
}
var
  FormatSettings: TFormatSettings;
begin
  FormatSettings.ShortDateFormat := 'yyyy-mm-dd';
  FormatSettings.DateSeparator := '-';
  FormatSettings.TimeSeparator := ':';
  if Length(S) = 7 then
    Result := SysUtils.StrToDate (S + '-01', FormatSettings)
  else
    Result := SysUtils.StrToDate (Copy(S,1,10), FormatSettings);
  if Length(S) in [16,19] then
    Result := Result + StrToTime(Copy(S,12,8));
end;

procedure TfrmMain.mniAmazonReturnsClick(Sender: TObject);
// "114-3942351-8041064","14otsHt+ShO8MPvjRYJGOw","2023-09-21T19:40:17.040Z","USD","18.35","Completed","Refund"

  procedure ProcessLine(S: string);
  var
    ListItem: TListItem;
    Trans: TChkTrans;
  begin
    if Length(S) = 0 then
      Exit;
    ParseLine(S);
    Trans := TChkTrans.Create;
    with Trans do begin
      OrderID := Fields[0];
      Date := ISOStrToDateTime(Fields[2]);
      AmountDebit := StrToAmt(Fields[4]);
      with lvPayments do begin
        ListItem := Items.Add;
        ListItem.Data := Trans;
        CopyAmznOrderToListItem(Trans, ListItem);
      end;
    end;
  end;

var
  Filename: string;
  I: Integer;
begin
  OpenDialog1.Filter := '*.csv|*.csv';
  OpenDialog1.Filename := 'Retail.OrdersReturned.Payments.1.csv';
  if OpenDialog1.Execute then begin
    with lvPayments do begin
      Columns[2].Caption := 'Order ID';
      Columns[3].Caption := 'Product Name';
    end;
    Filename := OpenDialog1.Filename;
    if not ClearListView(lvPayments) then
      Exit;
    lblFilenamePymt.Caption := Filename;
    slFileText.LoadFromFile(Filename);
    Delimiter := ',';
    CheckFileFormat('"OrderID","ReversalID","RefundCompletionDate","Currency","AmountRefunded","Status","DisbursementType"',
      slFileText[0]);
    try
      LoadingFile := True;
      for I := 1 to slFileText.Count-1 do
        ProcessLine(slFileText[I]);
    finally
      LoadingFile := False;
    end;
  end
end;

procedure TfrmMain.btnApplyClick(Sender: TObject);
var
  CT: TChkTrans;
  L: TListItem;
  I, N: Integer;
begin
  with lvChecking do begin
    if (SelCount > 0) and (cboCategory.ItemIndex <> -1) then begin
      I := SelCount;
      N := Items.IndexOf(Selected);
      while I > 0 do begin
        L := Items[N];
        if L.Selected then begin
          CT := TChkTrans(L.Data);
          CT.Category := cboCategory.Text;
          CopyChkTransToListItem(CT, L);
          Dec(I);
        end;
        Inc(N);
      end;
    end;
  end;
end;

procedure TfrmMain.FormShow(Sender: TObject);
begin
  cboCategory.Items.LoadFromFile('categories.txt');
end;

procedure TfrmMain.mniChaseVisaClick(Sender: TObject);
begin
  ReadCSVFile('Transaction Date,Post Date,Description,Category,Type,Amount,Memo',
    procedure (S: string; Trans: TChkTrans)
    {
    05/18/2020,05/18/2020,AUTOMATIC PAYMENT - THANK,,Payment,119.77,
    05/15/2020,05/15/2020,AMZN Mktp US*MC3CM8DJ2,Shopping,Sale,-9.13,
    }
    begin
      ParseLine(S);
      with Trans do begin
        Date := StrToDate(Fields[0]);
        Description := Fields[2];
        Category := Fields[3];
        AmountDebit := StrToAmt(Fields[5]);
        if AmountDebit > 0 then begin
          AmountCredit := AmountDebit;
          AmountDebit := 0;
        end;
        Memo := Fields[6];
        OrderID := Memo;
      end;
    end
  );
end;

procedure TfrmMain.mniCitiVisaClick(Sender: TObject);
begin
  ReadCSVFile('Status,Date,Description,Debit,Credit,Member Name',
    procedure (S: string; Trans: TChkTrans)
    {
    Cleared,09/13/2019,"AUTOPAY 000000000084386RAUTOPAY AUTO-PMT",,-3527.99,RONALD J SCHUSTER
    Cleared,09/13/2019,"OLIVE GARDEN 00010736 MIDDLEBRG HTSOH",41.00,,ELISABETH R SCHUSTER
    }
    begin
      ParseLine(S);
      with Trans do begin
        Date := StrToDate(Fields[1]);
        Description := Fields[2];
        AmountDebit := -StrToAmt(Fields[3]);
        AmountCredit := -StrToAmt(Fields[4]);
      end;
    end
  );
end;

procedure TfrmMain.mniConvertMintcategoriestoQuickenClick(Sender: TObject);
var
  I: Integer;
  MintToQuickenCatgs: TDictionary<string,string>;
  S, MintCatg, QuickenCatg: string;
  CT: TChkTrans;
  L: TListItem;
begin
  MintToQuickenCatgs := TDictionary<string,string>.Create;
  try
    slFileText.LoadFromFile('MintToQuickenCategories.csv');
    Delimiter := ',';
    for I := 1 to slFileText.Count-1 do begin
      S := Trim(slFileText[I]);
      ParseLine(S);
      MintCatg := Fields[0];
      QuickenCatg := Fields[1];
      MintToQuickenCatgs.Add(Uppercase(MintCatg), QuickenCatg);
    end;
    with lvChecking do
      for I := 0 to Items.Count-1 do begin
        L := Items[I];
        CT := TChkTrans(L.Data);
        if MintToQuickenCatgs.TryGetValue(Uppercase(CT.Category), QuickenCatg) then begin
          CT.Category := QuickenCatg;
          CopyChkTransToListItem(CT, L);
        end;
      end;
  finally
    MintToQuickenCatgs.Free;
  end;
end;

(*
procedure TForm1.mniAmazonOrdersClick(Sender: TObject);

  procedure ProcessLine(AOrderID, AProductName: string;
    AOrderDate: TDateTime; AAmount: Currency);
  var
    ListItem: TListItem;
    Trans: TChkTrans;
  begin
    Trans := TChkTrans.Create;
    with Trans do begin
      Memo := AOrderID;
      Description := AProductName;
      Date := AOrderDate;
      AmountDebit := -AAmount;
      with lvPayments do begin
        ListItem := Items.Add;
        ListItem.Data := Trans;
        CopyAmznOrderToListItem(Trans, ListItem);
      end;
    end;
  end;

var
  Filename: string;
  data: TBytes;
  JSONValue, jv: TJSONValue;
  joName: TJSONObject;
  OrderID, ProductName: string;
  OrderDate: TDateTime;
  Amount: Currency;
begin
  OpenDialog1.Filter := 'JSON files (*.json)|*.json';
  if OpenDialog1.Execute then begin
    Filename := OpenDialog1.Filename;
    if not ClearListView(lvPayments) then
      Exit;
    lblFilename.Caption := Filename;

    try
      LoadingFile := True;
      data := TEncoding.ASCII.GetBytes(TFile.ReadAllText(Filename));
      JSONValue := TJSONObject.ParseJSONValue(data, 0);
      for jv in JSONValue as TJSONArray do begin  // Returns TJSONValue
        joName := jv as TJSONObject;
        OrderID := joName.Get('Order ID').JSONValue.Value;
        OrderDate := joName.Get('Order Date').JSONValue.AsType<TDateTime>;
        Amount := joName.Get('Total Owed').JSONValue.AsType<Currency>;
        ProductName := joName.Get('Product Name').JSONValue.Value;
        ProcessLine(OrderID, ProductName, OrderDate, Amount)
      end{for};
    finally
      LoadingFile := False;
    end;
  end
end;
*)

procedure TfrmMain.mniAmazonOrdersClick(Sender: TObject);
// "Amazon.com","113-1463026-7469032","2023-12-07T23:50:08Z","Not Applicable","USD","11.99","0.91","0","'-0.6'","12.3","11.99","0.91","B0BCW59XZQ","New","1","Visa - 1732","Closed","Shipped","2023-12-08T16:39:37Z","next-1dc","Ronald J Schuster 18343 RIVER VALLEY BLVD NORTH ROYALTON OH 44133-6096 United States","Ronald J Schuster 18343 RIVER VALLEY BLVD NORTH ROYALTON OH 44133-6096 United States","AMZN_US(TBA310317999211)","BENTOBEN iPhone 14 Pro Max Case, Slim Lightweight 360° Ring Holder Kickstand Support Car Mount Shockproof Women Men Non-Slip Protective Case for iPhone 14 Pro Max 6.7"", White","Not Available","Not Available","Not Available"

  procedure ProcessLine(S: string);
  var
    ListItem: TListItem;
    Trans: TChkTrans;
  begin
    if Length(S) = 0 then
      Exit;
    ParseLine(S);
    Trans := TChkTrans.Create;
    with Trans do begin
      OrderID := Fields[1];
      Date := ISOStrToDateTime(Fields[2]);
      AmountDebit := StrToAmt(Fields[9]);
      Description := Fields[23];
      with lvPayments do begin
        ListItem := Items.Add;
        ListItem.Data := Trans;
        CopyAmznOrderToListItem(Trans, ListItem);
      end;
    end;
  end;

var
  Filename: string;
  I: Integer;
begin
  OpenDialog1.Filter := '*.csv|*.csv';
  OpenDialog1.Filename := 'Retail.OrderHistory.1.csv';
  if OpenDialog1.Execute then begin
    with lvPayments do begin
      Columns[2].Caption := 'Order ID';
      Columns[3].Caption := 'Product Name';
    end;
    Filename := OpenDialog1.Filename;
    if not ClearListView(lvPayments) then
      Exit;
    lblFilenamePymt.Caption := Filename;
    slFileText.LoadFromFile(Filename);
    Delimiter := ',';
    CheckFileFormat('"Website","Order ID","Order Date","Purchase Order Number","Currency","Unit Price","Unit Price Tax","Shipping Charge","Total Discounts","Total Owed","Shipment Item Subtotal","Shipment Item Subtotal Tax","ASIN","Product Condition","Quantity",'+
      '"Payment Instrument Type","Order Status","Shipment Status","Ship Date","Shipping Option","Shipping Address","Billing Address","Carrier Name & Tracking Number","Product Name","Gift Message","Gift Sender Name","Gift Recipient Contact Details",'+
      '"Item Serial Number"',
      slFileText[0]);
    try
      LoadingFile := True;
      for I := 1 to slFileText.Count-1 do
        ProcessLine(slFileText[I]);
    finally
      LoadingFile := False;
    end;
  end
end;

procedure TfrmMain.mniChaseCheckingClick(Sender: TObject);
begin
  ReadCSVFile('Details,Posting Date,Description,Amount,Type,Balance,Check or Slip #',
    procedure (S: string; Trans: TChkTrans)
    {
    DEBIT,12/09/2016,"HI ROCKY SPORTS INC BUENA VISTA CO           12/08",-73.35,DEBIT_CARD,1741.09,,
    }
    begin
      ParseLine(S);
      with Trans do begin
        Date := StrToDate(Fields[1]);
        Description := Fields[2];
        AmountDebit := StrToAmt(Fields[3]);
        if AmountDebit > 0 then begin
          AmountCredit := AmountDebit;
          AmountDebit := 0;
        end;
      end;
    end
  );
end;

procedure TfrmMain.mniMintClick(Sender: TObject);
begin

  ReadCSVFile('"Date","Description","Original Description","Amount","Transaction Type","Category","Account Name","Labels","Notes"',
    procedure (S: string; Trans: TChkTrans)
    {
    "2/15/2022","Deposit from EJ for payment to TLC","Deposit from EJ for payment to","500.00","credit","Transfer","Interest Checking","",""
    "2/15/2022","CITI AUTOPAY - PAYMENT","External Withdrawal CITI","3242.52","debit","Transfer","Interest Checking","",""

    }
    begin
      ParseLine(S);
      with Trans do begin
        Date := StrToDate(Fields[0]);
        Description := Fields[1];
        Memo := Fields[2];
        AmountDebit := StrToAmt(Fields[3]);
        if Fields[4] = 'debit' then
          AmountDebit := -AmountDebit
        else begin
          AmountCredit := AmountDebit;
          AmountDebit := 0;
        end;
        Category := Fields[5];
      end;
    end
  );
end;

procedure TfrmMain.mniProcessAmazonTransactionsClick(Sender: TObject);
var
  sOrderID: string;
  ListItem: TListItem;
  TransList, OrderList: TList<TChkTrans>;
  T, I, K: Integer;
  SelectedOrders: TList<Integer>;
  Total: Currency;
  OrdsPerLine: array [0..1] of Integer;
begin
  TransList := nil;
  OrderList := nil;
  SelectedOrders := nil;
  try
    TransList := TList<TChkTrans>.Create;
    OrderList := TList<TChkTrans>.Create;
    SelectedOrders := TList<Integer>.Create;
    for T := 0 to lvChecking.Items.Count - 1 do begin
      sOrderID :=  TChkTrans(lvChecking.Items[T].Data).OrderID;
      if not TChkTrans(lvChecking.Items[T].Data).AmznTransProcessed and (sOrderID <> '') then begin
        TransList.Clear;
        OrderList.Clear;
        for I := 0 to lvChecking.Items.Count - 1 do
          if TChkTrans(lvChecking.Items[I].Data).OrderID = sOrderID then
            TransList.Add(TChkTrans(lvChecking.Items[I].Data));
        for I := 0 to lvPayments.Items.Count - 1 do
          with TChkTrans(lvPayments.Items[I].Data) do
            if (OrderID = sOrderID) and (AmountDebit <> 0) then
              OrderList.Add(TChkTrans(lvPayments.Items[I].Data));
        for I := 0 to TransList.Count - 1 do begin
          for var J := 1 to Round(Power(2, OrderList.Count)) - 1 do begin
            SelectedOrders.Clear;
            var A := J;
            var B := 0;
            while A > 0 do begin
              if Odd(A) then
                SelectedOrders.Add(B);
              A := A shr 1;
              Inc(B);
            end;
            Total := 0;
            for K := 0 to SelectedOrders.Count - 1 do
              Total := Total + OrderList[SelectedOrders[K]].AmountDebit;
            if SameValue(Total, -1 * TransList[I].AmountDebit)
            or SameValue(Total, TransList[I].AmountCredit) then begin
              with TransList[I] do begin
                if SelectedOrders.Count < 2 then
                  OrdsPerLine[0] := SelectedOrders.Count
                else
                  OrdsPerLine[0] := SelectedOrders.Count div 2;
                OrdsPerLine[1] := SelectedOrders.Count - OrdsPerLine[0];
                Description := 'AMZN';
                var CharsAvail := 59 - (OrdsPerLine[0] * 2);
                for K := 0 to OrdsPerLine[0] - 1 do
                  Description := Description + IfThen(K = 0, ': ', '; ') + LeftStr(OrderList[SelectedOrders[K]].Description,
                    CharsAvail div OrdsPerLine[0]);
                CharsAvail := 64 - Length(Memo) - (OrdsPerLine[1] * 2);
                for K := OrdsPerLine[0] to SelectedOrders.Count - 1 do
                  Memo := Memo + IfThen(K > 0, '; ') + LeftStr(OrderList[SelectedOrders[K]].Description,
                    CharsAvail div OrdsPerLine[1]);
                AmznTransProcessed := True;
              end;
            end;
          end;
        end;
        CopyChkTransToListItem(TChkTrans(lvChecking.Items[T].Data), lvChecking.Items[T]);
      end;
    end;
  finally
    TransList.Free;
    OrderList.Free;
    SelectedOrders.Free;
  end;
end;

{ TChkTrans }

function TChkTrans.ValueByName(Name: string): Variant;
begin
  var I := IndexText(Name,
   [{0}     'Trans No',
    {1}     'Date',
    {2,3,4} 'Payee', 'Desc', 'Product Name',
    {5}     'Memo',
    {6,7}   'Amount', 'Debit',
    {8}     'Credit',
    {9}     'Balance',
    {10}    'Check No',
    {11}    'Fees',
    {12}    'Category',
    {13}    'Order ID']);
  case I of
    0:     Result := TransactionNumber;
    1:     Result := Date;
    2,3,4: Result := UpperCase(Description);
    5:     Result := Memo;
    6,7:   Result := AmountDebit;
    8:     Result := AmountCredit;
    9:     Result := Balance;
    10:    Result := CheckNumber;
    11:    Result := Fees;
    12:    Result := Category;
    13:    Result := OrderID;
    else   Result := '';
  end;
end;

end.
