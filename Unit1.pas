unit Unit1;
{
01/24/10   RJS   Refactored btnReadPaymentClick: pulled QIF reading code out
                 into new procedure ReadQIFFile.
05/13/11   RJS   Added support for Category.
}


interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ComCtrls, ExtCtrls, Mask, wwdbedit, Wwdotdot, Wwdbcomb;

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

  TForm1 = class(TForm)
    OpenDialog1: TOpenDialog;
    Panel1: TPanel;
    btnReadChecking: TButton;
    pnlChecking: TPanel;
    lvChecking: TListView;
    btnCleanup: TButton;
    btnReadPayment: TButton;
    Splitter1: TSplitter;
    Panel3: TPanel;
    lvPayments: TListView;
    Timer1: TTimer;
    btnWriteQIF: TButton;
    SaveDialog1: TSaveDialog;
    btnReadCheckbook: TButton;
    cboCategory: TwwDBComboBox;
    Label1: TLabel;
    btnApply: TButton;
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
  private
    { Private declarations }
    slFileText: TStringList;
    Delimiter: Char;
    procedure CleanUpItem (Item: TListItem);
    procedure CopyChkTransToListItem(T:TChkTrans; L:TListItem);
    procedure CopyPmtTransToListItem(T:TChkTrans; L:TListItem);
    procedure ReadQIFFile(Filename: string; ListView: TListView;
      CopyTransToListItem: TCopyTransToListItem);
    procedure ClearListView(LV: TListView);
    function GetField(InStr: string; var Pos: Integer): string;
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.DFM}

uses
  StStrL;

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
  Phrase: array[1..4] of string = (
    'External Deposit',
    'External Withdrawal',
    'Point Of Sale Deposit',
    'Point Of Sale Withdrawal');
begin
{ For all of these phrases, if phrase found at beginning of Desc, delete the phrase
  from Desc. If remaining Desc is not blank, add a space. Append Memo to end of Desc }
  Trans := TChkTrans(Item.Data);
  Desc := Trans.Description;
  Memo := Trans.Memo;

  for I := 1 to 4 do
    if Pos(Phrase[I], Desc) = 1 then begin
      Delete(Desc, 1, Length(Phrase[I])+1);
      if Desc <> '' then
        Desc := Desc + ' ';
      Trans.Description := Desc + Memo;
      Trans.Memo := '';
      Break;
    end;

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
  if S = '' then
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

procedure TForm1.ClearListView(LV: TListView);
var
  I: Integer;
begin
  for I := 0 to LV.Items.Count-1 do
    TChkTrans(LV.Items[I]).Free;
  LV.Items.Clear;
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
begin
  OpenDialog1.Filter := '*.csv|*.csv|*.qif|*.qif';
  if OpenDialog1.Execute then begin
    Filename := OpenDialog1.Filename;
    if SameText(ExtractFileEXt(Filename), '.qif') then
      ReadQIFFile(Filename, lvChecking, CopyChkTransToListItem)
    else begin
      ClearListView(lvChecking);
      slFileText.LoadFromFile(Filename);
      Delimiter := ',';
      if slFileText[5] <> 'Transaction Number,Date,Description,Memo,Amount Debit,Amount Credit,Balance,Check Number,Fees' then
        raise Exception.Create('Invalid file input format');
      for I := 6 to slFileText.Count-1 do
        ProcessLine(slFileText[I]);
    end
  end;
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
  ClearListView(ListView);
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
          'N': if S[2] in ['0'..'9'] then
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
// "12/31/2010","55.00",,"Uncategorized","PMT","FIVE STAR GYMNASTICS"

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
      Date := StrToDate(GetField(S,Pos));
      AmountDebit := -StrToAmt(GetField(S,Pos));
      GetField(S,Pos);
      Memo := GetField(S,Pos);
      GetField(S,Pos);
      Description := GetField(S,Pos);

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
  OpenDialog1.Filter := '*.csv|*.csv';
  if OpenDialog1.Execute then begin
    Filename := OpenDialog1.Filename;
    ClearListView(lvPayments);
    slFileText.LoadFromFile(Filename);
    Delimiter := ',';
    for I := 1 to slFileText.Count-1 do
      ProcessLine(slFileText[I]);
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
      ClearListView(lvPayments);
      slFileText.LoadFromFile(Filename);
      Delimiter := #9;
      if slFileText[0] <> 'Chk#'#9'Date'#9'Payee'#9'Memo'#9'Amt' then
        raise Exception.Create('Invalid file input format');
      for I := 1 to slFileText.Count-1 do
        ProcessLine(slFileText[I]);
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

procedure TForm1.lvCheckingDragOver(Sender, Source: TObject; X, Y: Integer;
  State: TDragState; var Accept: Boolean);
var
  CT: TChkTrans;
  L: TListItem;
begin
  L := lvChecking.GetItemAt(X,Y);
  if (L = nil) or (DragPmt = nil) then
    Exit;
  CT := TChkTrans(L.Data);
  Accept := (CT.AmountDebit = DragPmt.AmountDebit)
             and (CT.Date >= DragPmt.Date);
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

end.
