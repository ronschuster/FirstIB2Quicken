unit Unit1;
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
  end;

  TPmtTrans = class
    Date: TDateTime;
    Amt: Double;
    Payee: string;
    CheckNumber: Integer;
  end;

  TCopyTransToListItem = procedure (T:TChkTrans; L:TListItem) of Object;
  TProcessLineProc = reference to procedure (S: string; Trans: TChkTrans);

  TForm1 = class(TForm)
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
    mniChaseChecking: TMenuItem;
    Panel2: TPanel;
    Label1: TLabel;
    cboCategory: TwwDBComboBox;
    btnApply: TButton;
    mniChaseVisa: TMenuItem;
    mniCitiVisa: TMenuItem;
    lblFilename: TLabel;
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
  private
    { Private declarations }
    slFileText: TStringList;
    Delimiter: Char;
    LoadingFile: Boolean;  // while loading file in the right-hand pane, suppress custom draw on left-hand pane
    procedure CleanUpItem (Item: TListItem);
    procedure CopyChkTransToListItem(T:TChkTrans; L:TListItem);
    procedure CopyPmtTransToListItem(T:TChkTrans; L:TListItem);
    procedure ReadQIFFile(Filename: string; ListView: TListView;
      CopyTransToListItem: TCopyTransToListItem);
    function ClearListView(LV: TListView): Boolean;
    function GetField(InStr: string; var Pos: Integer): string;
    function MatchingTransaction(CT, Pmt: TChkTrans): Boolean;
    procedure ReadCSVFile(HeaderLine: string; AProcessLine: TProcessLineProc);
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.DFM}

uses
  StStrL, RegularExpressions;

var
  ChkSortField, PmtSortField: string;
  ChkSortAsc, PmtSortAsc: Boolean;
  DragPmt: TChkTrans;
  DragIndex: Integer;


procedure TForm1.CleanUpItem (Item: TListItem);
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

procedure TForm1.CopyChkTransToListItem(T:TChkTrans; L:TListItem);
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

procedure TForm1.CopyPmtTransToListItem(T:TChkTrans; L:TListItem);
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

function TForm1.ClearListView(LV: TListView): Boolean;
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

procedure TForm1.btnReadCheckingClick(Sender: TObject);

  procedure ProcessLine(S: string);
  var
    Pos: Integer;
    ListItem: TListItem;
    Trans: TChkTrans;
  begin
    if Length(S) = 0 then
      Exit;

    Pos := 1;
    Trans := TChkTrans.Create;
    with Trans do begin
      TransactionNumber := GetField(S,Pos);
      Date := StrToDate(GetField(S,Pos));
      Description := GetField(S,Pos);
      Memo := GetField(S,Pos);
      AmountDebit := StrToAmt(GetField(S,Pos));
      AmountCredit := StrToAmt(GetField(S,Pos));
      Balance := StrToAmt(GetField(S,Pos));
      CheckNumber := StrToIntDef(GetField(S,Pos),0);
      Fees := StrToAmt(GetField(S,Pos));

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
    lblFilename.Caption := Filename;
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



procedure TForm1.btnReadDiscoverFileClick(Sender: TObject);
begin
  ReadCSVFile('Trans. Date,Post Date,Description,Amount,Category',
    procedure (S: string; Trans: TChkTrans)
    var
      Pos: Integer;
    begin
      Pos := 1;
      with Trans do begin
        Date := StrToDate(GetField(S,Pos));
        GetField(S,Pos); //Post Date
        Description := GetField(S,Pos);
        AmountDebit := -1 * StrToAmt(GetField(S,Pos));
        if AmountDebit > 0 then begin
          AmountCredit := AmountDebit;
          AmountDebit := 0;
        end;
        Category := GetField(S,Pos);
      end;
    end
  );
end;

procedure TForm1.lvCheckingCompare(Sender: TObject; Item1, Item2: TListItem;
  Data: Integer; var Compare: Integer);
var
  Trans1, Trans2: TChkTrans;
  Field1, Field2: Variant;
begin
  Trans1 := TChkTrans(Item1.Data);
  Trans2 := TChkTrans(Item2.Data);
  if ChkSortField = 'Trans No' then begin
    Field1 := Trans1.TransactionNumber;
    Field2 := Trans2.TransactionNumber;
  end
  else if ChkSortField = 'Date' then begin
    Field1 := Trans1.Date;
    Field2 := Trans2.Date;
  end
  else if ChkSortField = 'Desc' then begin
    Field1 := Uppercase(Trans1.Description);
    Field2 := Uppercase(Trans2.Description);
  end
  else if ChkSortField = 'Memo' then begin
    Field1 := Uppercase(Trans1.Memo);
    Field2 := Uppercase(Trans2.Memo);
  end
  else if ChkSortField = 'Debit' then begin
    Field1 := Trans1.AmountDebit;
    Field2 := Trans2.AmountDebit;
  end
  else if ChkSortField = 'Credit' then begin
    Field1 := Trans1.AmountCredit;
    Field2 := Trans2.AmountCredit;
  end
  else if ChkSortField = 'Balance' then begin
    Field1 := Trans1.Balance;
    Field2 := Trans2.Balance;
  end
  else if ChkSortField = 'Category' then begin
    Field1 := Trans1.Category;
    Field2 := Trans2.Category;
  end
  else if ChkSortField = 'Check No' then begin
    Field1 := Trans1.CheckNumber;
    Field2 := Trans2.CheckNumber;
  end
  else if ChkSortField = 'Fees' then begin
    Field1 := Trans1.Fees;
    Field2 := Trans2.Fees;
  end;
  if Field1= Field2 then
    Compare := 0
  else if Field1 < Field2 then
    Compare := -1
  else
    Compare := 1;
  if not ChkSortAsc then
    Compare := -1 * Compare;
end;

procedure TForm1.lvCheckingCustomDrawItem(Sender: TCustomListView;
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

procedure TForm1.btnCleanupClick(Sender: TObject);
var
  I: Integer;
begin
  for I := 0 to lvChecking.Items.Count-1 do
    CleanUpItem(lvChecking.Items[I]);
end;

procedure TForm1.lvCheckingColumnClick(Sender: TObject;
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

procedure TForm1.ReadQIFFile(Filename: string; ListView: TListView; CopyTransToListItem: TCopyTransToListItem);
var
  I, J, NbrTrans: Integer;
  Trans: TChkTrans;
  ListItem: TListItem;
  S: string;
begin
  if not ClearListView(ListView) then
    Exit;
  lblFilename.Caption := Filename;
  slFileText.LoadFromFile(Filename);
  if slFileText[0] <> '!Type:Bank' then
    raise Exception.Create('Invalid file input format');

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

function TForm1.GetField(InStr: string; var Pos: Integer): string;
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

procedure TForm1.btnReadPaymentClick(Sender: TObject);
// 12/31/2015,"Grace C & MA Church",$562.00,JX1F0-1RGX0,DDA - 2206,Complete,Charitable Donations

  procedure ProcessLine(S: string);
  var
    Pos: Integer;
    ListItem: TListItem;
    Trans: TChkTrans;
    Status: string;
  begin
    if Length(S) = 0 then
      Exit;
    Pos := 1;
    Trans := TChkTrans.Create;
    with Trans do begin
      Date := StrToDate(GetField(S,Pos));
      Description := GetField(S,Pos);
      AmountDebit := -StrToAmt(FilterL(GetField(S,Pos),'$'));
      GetField(S,Pos);
      GetField(S,Pos);
      Status := GetField(S,Pos);
      Category := GetField(S,Pos);
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
    Filename := OpenDialog1.Filename;
    if not ClearListView(lvPayments) then
      Exit;
    lblFilename.Caption := Filename;
    slFileText.LoadFromFile(Filename);
    Delimiter := ',';
    if slFileText[0] <> 'Deliver by,Paid to,Amount,Confirmation No.,Paid from,Status,Category' then
      raise Exception.Create('Invalid file input format');
    try
      LoadingFile := True;
      for I := 1 to slFileText.Count-1 do
        ProcessLine(slFileText[I]);
    finally
      LoadingFile := False;
    end;
  end
end;

procedure TForm1.btnReadCheckbookClick(Sender: TObject);
  procedure ProcessLine(S: string);
  var
    Pos: Integer;
    ListItem: TListItem;
    Trans: TChkTrans;
  begin
    if Length(S) = 0 then
      Exit;
    Pos := 1;
    Trans := TChkTrans.Create;
    with Trans do begin
      CheckNumber := StrToIntDef(GetField(S,Pos),0);
      Date := StrToDate(GetField(S,Pos));
      Description := GetField(S,Pos);
      Memo := GetField(S,Pos);
      AmountDebit := -StrToAmt(GetField(S,Pos));

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
    Filename := OpenDialog1.Filename;
    if SameText(ExtractFileEXt(Filename), '.qif') then
      ReadQIFFile(Filename, lvPayments, CopyPmtTransToListItem)
    else begin
      if not ClearListView(lvPayments) then
        Exit;
      lblFilename.Caption := Filename;
      slFileText.LoadFromFile(Filename);
      Delimiter := #9;
      if slFileText[0] <> 'Chk#'#9'Date'#9'Payee'#9'Memo'#9'Amt' then
        raise Exception.Create('Invalid file input format');
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

procedure TForm1.FormCreate(Sender: TObject);
begin
  slFileText := TStringList.Create;
end;

procedure TForm1.FormDestroy(Sender: TObject);
begin
  slFileText.Free;
end;

procedure TForm1.lvPaymentsChange(Sender: TObject; Item: TListItem;
  Change: TItemChange);
begin
  if not LoadingFile then
    lvChecking.Repaint;
end;

procedure TForm1.lvPaymentsColumnClick(Sender: TObject;
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

procedure TForm1.lvPaymentsCompare(Sender: TObject; Item1,
  Item2: TListItem; Data: Integer; var Compare: Integer);
var
  Trans1, Trans2: TChkTrans;
  Field1, Field2: Variant;
begin
  Trans1 := TChkTrans(Item1.Data);
  Trans2 := TChkTrans(Item2.Data);
  if PmtSortField = 'Date' then begin
    Field1 := Trans1.Date;
    Field2 := Trans2.Date;
  end
  else if PmtSortField = 'Amount' then begin
    Field1 := Trans1.AmountDebit;
    Field2 := Trans2.AmountDebit;
  end
  else if PmtSortField = 'Payee' then begin
    Field1 := UpperCase(Trans1.Description);
    Field2 := UpperCase(Trans2.Description);
  end;
  if Field1= Field2 then
    Compare := 0
  else if Field1 < Field2 then
    Compare := -1
  else
    Compare := 1;
  if not PmtSortAsc then
    Compare := -1 * Compare;
end;

procedure TForm1.lvPaymentsKeyDown(Sender: TObject; var Key: Word;
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

function TForm1.MatchingTransaction(CT, Pmt: TChkTrans): Boolean;
{ returns true if the checking transaction CT matches the payment Pmt for the
  purposes of highlighting valid drag targets. }
begin
  Result := (CT.AmountDebit = Pmt.AmountDebit)
             and (CT.Date >= Pmt.Date);
end;

procedure TForm1.ReadCSVFile(HeaderLine: string; AProcessLine: TProcessLineProc);

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
    lblFilename.Caption := Filename;
    slFileText.LoadFromFile(Filename);
    Delimiter := ',';
    while Trim(slFileText[0]) = '' do
      slFileText.Delete(0);
    if Trim(slFileText[0]) <> HeaderLine then
      raise Exception.Create('Invalid file input format.'#13'Expected: ' + HeaderLine + #13'Received: ' + Trim(slFileText[0]));
    for I := 1 to slFileText.Count-1 do
      ProcessLine(Trim(slFileText[I]));
  end
end;

procedure TForm1.mniChaseCheckingClick(Sender: TObject);
begin
  ReadCSVFile('Details,Posting Date,Description,Amount,Type,Balance,Check or Slip #',
    procedure (S: string; Trans: TChkTrans)
    {
    DEBIT,12/09/2016,"HI ROCKY SPORTS INC BUENA VISTA CO           12/08",-73.35,DEBIT_CARD,1741.09,,
    }
    var
      Pos: Integer;
    begin
      Pos := 1;
      with Trans do begin
        GetField(S,Pos); //Details
        Date := StrToDate(GetField(S,Pos));
        Description := GetField(S,Pos);
        AmountDebit := StrToAmt(GetField(S,Pos));
        if AmountDebit > 0 then begin
          AmountCredit := AmountDebit;
          AmountDebit := 0;
        end;
      end;
    end
  );
end;

procedure TForm1.mniExitClick(Sender: TObject);
begin
  Close
end;

procedure TForm1.lvCheckingDragOver(Sender, Source: TObject; X, Y: Integer;
  State: TDragState; var Accept: Boolean);
var
  L: TListItem;
begin
  L := lvChecking.GetItemAt(X,Y);
  if (L = nil) or (DragPmt = nil) then
    Exit;
  Accept := MatchingTransaction(TChkTrans(L.Data), DragPmt);
end;

procedure TForm1.lvPaymentsStartDrag(Sender: TObject;
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

procedure TForm1.lvCheckingDragDrop(Sender, Source: TObject; X,
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

procedure TForm1.Timer1Timer(Sender: TObject);
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

procedure TForm1.pnlCheckingDragOver(Sender, Source: TObject; X,
  Y: Integer; State: TDragState; var Accept: Boolean);
begin
  Accept := False;
  Timer1.Enabled := True;
end;

procedure TForm1.btnWriteQIFClick(Sender: TObject);
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

procedure TForm1.btnReadFinanceorksFileClick(Sender: TObject);
begin
  ReadCSVFile('Date, Account Name, Check #, Transaction, Category, Note, Expense, Deposit',
    procedure (S: string; Trans: TChkTrans)
    var
      Pos: Integer;
    begin
      Pos := 1;
      with Trans do begin
        Date := StrToDate(GetField(S,Pos));
        GetField(S,Pos);
        CheckNumber :=  StrToIntDef(GetField(S,Pos),0);
        Description := GetField(S,Pos);
        Category := GetField(S,Pos);
        Memo := GetField(S,Pos);
        AmountDebit := StrToAmt(GetField(S,Pos));
        AmountCredit := StrToAmt(GetField(S,Pos));
      end;
    end
  );
end;

procedure TForm1.btnApplyClick(Sender: TObject);
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

procedure TForm1.FormShow(Sender: TObject);
begin
  cboCategory.Items.LoadFromFile('categories.txt');
end;

procedure TForm1.mniChaseVisaClick(Sender: TObject);
begin
  ReadCSVFile('Transaction Date,Post Date,Description,Category,Type,Amount,Memo',
    procedure (S: string; Trans: TChkTrans)
    {
    05/18/2020,05/18/2020,AUTOMATIC PAYMENT - THANK,,Payment,119.77,
    05/15/2020,05/15/2020,AMZN Mktp US*MC3CM8DJ2,Shopping,Sale,-9.13,
    }
    var
      Pos: Integer;
    begin
      Pos := 1;
      with Trans do begin
        Date := StrToDate(GetField(S,Pos));
        GetField(S,Pos); //Post Date
        Description := GetField(S,Pos);
        Category := GetField(S,Pos);
        GetField(S,Pos); //Type
        AmountDebit := StrToAmt(GetField(S,Pos));
        if AmountDebit > 0 then begin
          AmountCredit := AmountDebit;
          AmountDebit := 0;
        end;
        Memo := GetField(S,Pos);
      end;
    end
  );
end;

procedure TForm1.mniCitiVisaClick(Sender: TObject);
begin
  ReadCSVFile('Status,Date,Description,Debit,Credit,Member Name',
    procedure (S: string; Trans: TChkTrans)
    {
    Cleared,09/13/2019,"AUTOPAY 000000000084386RAUTOPAY AUTO-PMT",,-3527.99,RONALD J SCHUSTER
    Cleared,09/13/2019,"OLIVE GARDEN 00010736 MIDDLEBRG HTSOH",41.00,,ELISABETH R SCHUSTER
    }
    var
      Pos: Integer;
    begin
      Pos := 1;
      with Trans do begin
        GetField(S,Pos); //Status
        Date := StrToDate(GetField(S,Pos));
        Description := GetField(S,Pos);
        AmountDebit := -StrToAmt(GetField(S,Pos));
        AmountCredit := -StrToAmt(GetField(S,Pos));
      end;
    end
  );
end;

end.
