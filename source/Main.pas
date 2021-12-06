
unit
  Main;

// --------------------------------------------------------------------------
//
// --------------------------------------------------------------------------

interface

// --------------------------------------------------------------------------
//
// --------------------------------------------------------------------------

uses
  Windows,
  Messages,
  SysUtils,
  Classes,
  Graphics,
  Controls,
  Forms,
  Dialogs,
  ExtCtrls,
  ComCtrls,
  StdCtrls;

// --------------------------------------------------------------------------
//
// --------------------------------------------------------------------------

type
  PFileInfo = ^TFileInfo;
  TFileInfo = packed record
    Directory: Boolean;
    CreationTime: TDateTime;
    LastModifiedTime: TDateTime;
    LastAccessedTime: TDateTime;
    Size: Integer;
  end;

// --------------------------------------------------------------------------
//
// --------------------------------------------------------------------------

type
  TfrmStartup = class(TForm)

    gbOptions: TGroupBox;
    lbFolder: TLabel;
    edFolder: TEdit;
    btnChoose: TButton;
    lbStyle: TLabel;
    edStyle: TComboBox;
    gbFilters: TGroupBox;
    lbNameContains: TLabel;
    edNameContains: TEdit;
    lbNameContainsPrototypes: TLabel;
    edNameContainsPrototypes: TComboBox;
    lbSmallerThan: TLabel;
    edSmallerThan: TEdit;
    lbGreaterThan: TLabel;
    edGreaterThan: TEdit;
    lbCreatedBefore: TLabel;
    edCreatedBefore: TDateTimePicker;
    lbCreatedAfter: TLabel;
    edCreatedAfter: TDateTimePicker;
    lbLastModifiedBefore: TLabel;
    edLastModifiedBefore: TDateTimePicker;
    edLastModifiedAfter: TDateTimePicker;
    lbLastModifiedAfter: TLabel;
    lbLastAccessedBefore: TLabel;
    edLastAccessedBefore: TDateTimePicker;
    lbLastAccessedAfter: TLabel;
    edLastAccessedAfter: TDateTimePicker;
    lbNameDoesNotContain: TLabel;
    edNameDoesNotContain: TEdit;
    lbNameDoesNotContainPrototypes: TLabel;
    edNameDoesNotContainPrototypes: TComboBox;
    btnLoadProfile: TButton;
    btnSaveProfileAs: TButton;
    sbInfo: TStatusBar;
    btnClearNameContains: TButton;
    btnClearNameDoesNotContain: TButton;
    dlgLoadProfile: TOpenDialog;
    dlgSaveProfileAs: TSaveDialog;
    btnBrowse: TButton;
    gbSorting: TGroupBox;
    rgSorting: TRadioGroup;
    rbSortingNone: TRadioButton;
    rbSortingByName: TRadioButton;
    rbSortingBySize: TRadioButton;
    rbSortingByCreationTime: TRadioButton;
    rbSortingByLastModifiedTime: TRadioButton;
    rbSortingByLastAccessedTime: TRadioButton;
    rbSortingScramble: TRadioButton;
    cbSortingReverseOrder: TCheckBox;
    gbCase: TGroupBox;
    rgCase: TRadioGroup;
    rbLowerCase: TRadioButton;
    rbUpperCase: TRadioButton;
    rbProperCase: TRadioButton;
    rbUnchangedCase: TRadioButton;
    cbHideEmptyFolders: TCheckBox;
    cbHideFiles: TCheckBox;
    cbEnabledCreated: TCheckBox;
    cbEnabledLastModified: TCheckBox;
    cbEnabledLastAccessed: TCheckBox;

    procedure FormCreate(Sender: TObject);
    procedure btnBrowseClick(Sender: TObject);
    procedure btnClearNameContainsClick(Sender: TObject);
    procedure btnClearNameDoesNotContainClick(Sender: TObject);
    procedure btnLoadProfileClick(Sender: TObject);
    procedure btnSaveProfileAsClick(Sender: TObject);
    procedure edNameContainsPrototypesChange(Sender: TObject);
    procedure edNameDoesNotContainPrototypesChange(Sender: TObject);
    procedure btnChooseClick(Sender: TObject);

  private

    // Application path
    AppPath: String;

    // Holds in memory the whole hierarchy of file info objects
    trTree: TTreeView;

    // All the search, sorting and filtering parameters
    pmStyle: Integer;
    pmFolder, pmDestination: String;
    pmNameContains, pmNameDoesNotContain: TStringList;
    pmSmallerThan, pmGreaterThan: Integer;
    pmCreatedBefore, pmCreatedAfter: TDateTime;
    pmLastModifiedBefore, pmLastModifiedAfter: TDateTime;
    pmLastAccessedBefore, pmLastAccessedAfter: TDateTime;
    pmEnabledCreated, pmEnabledLastModified, pmEnabledLastAccessed: Boolean;
    pmHideEmptyFolders: Boolean; pmHideFiles: Boolean;
    pmSortingNone, pmSortingByName, pmSortingBySize, pmSortingByCreationTime, pmSortingByLastModifiedTime, pmSortingByLastAccessedTime, pmSortingScramble, pmSortingReverseOrder: Boolean;
    pmLowerCase, pmUpperCase, pmProperCase, pmUnchangedCase: Boolean;

    // The contents of the output document
    lsOutput: TStringList;

    function  ApplyFilters(Path: String; Data: TSearchRec): Boolean;

    procedure LowerCaseTree(Daddy: TTreeNode);
    procedure UpperCaseTree(Daddy: TTreeNode);
    procedure ProperCaseTree(Daddy: TTreeNode);

    procedure SearchClear;
    procedure SearchClearNodeData(Daddy: TTreeNode);
    procedure SearchScan(Path: String; Daddy: TTreeNode);
    procedure SearchPostProcess;
    procedure SearchShowResults;
    procedure SearchSort;

    procedure ApplyStyleGeneric;
    procedure ApplyStyleGoodOldTree(Daddy: TTreeNode; Level: Integer);
    procedure ApplyStyleHtmlTreeBriefPrintable(Daddy: TTreeNode; Level: Integer);
    procedure ApplyStyleHtmlTreeDetailedPrintable(Daddy: TTreeNode; Level: Integer);
    procedure ApplyStyleHtmlTreeEverythingPrintable(Daddy: TTreeNode; Level: Integer);
    procedure ApplyStyleHtmlTreeDetailedMultimedia(Daddy: TTreeNode; Level: Integer);
    procedure ApplyStyleHtmlTreeCompactMultimedia(Daddy: TTreeNode; Level: Integer);
    procedure ApplyStyleWindowsMediaPlayerPlaylist(Daddy: TTreeNode);

    procedure LoadProfile(FileName: String);
    procedure SaveProfileAs(FileName: String);

  end;

// --------------------------------------------------------------------------
//
// --------------------------------------------------------------------------

var
  frmStartup: TfrmStartup;

// --------------------------------------------------------------------------
//
// --------------------------------------------------------------------------

implementation

// --------------------------------------------------------------------------
//
// --------------------------------------------------------------------------

uses
  ShellAPI,
  BrowseForFolder;

// --------------------------------------------------------------------------
//
// --------------------------------------------------------------------------

{$R *.DFM}

// --------------------------------------------------------------------------
//
// --------------------------------------------------------------------------

procedure Split(Str: String; Sep: Char; var Lst: TStringList);
var
  j: Integer;
  Tmp: String;
begin
  Lst.Clear; if (Length(Str) > 0) then begin
    repeat
      j := Pos(Sep, Str); if (j > 0) then begin
        Tmp := Trim(Copy(Str, 1, j - 1));
        if (Length(Tmp) > 0) then Lst.Append(Tmp); Delete(Str, 1, j);
      end;
    until (j = 0); Tmp := Trim(Str); if (Length(Tmp) > 0) then Lst.Append(Tmp);
  end;
end;

// --------------------------------------------------------------------------
//
// --------------------------------------------------------------------------

function LoCase(C: Char): Char;
begin
  if (C in ['A'..'Z']) then Result := Char(Byte(C) + 32) else Result := C;
end;

// --------------------------------------------------------------------------
//
// --------------------------------------------------------------------------

function ProperCase(Name: String): String;
var
  j: Integer;
begin
  if (Length(Name) > 0) then begin
    if (Name = LowerCase(Name)) or (Name = UpperCase(Name)) then begin
      Name[1] := UpCase(Name[1]); for j := 2 to Length(Name) do begin
        if (Name[j-1] = #32) then Name[j] := UpCase(Name[j]) else Name[j] := LoCase(Name[j]);
      end;
    end;
  end; Result := Name;
end;

// --------------------------------------------------------------------------
//
// --------------------------------------------------------------------------

function StringContainsListSubstr(Str: String; var Lst: TStringList): Boolean;
var
  j: Integer;
  Fnd: Boolean;
begin
  j := 0; Fnd := False; while (j < Lst.Count) and not(Fnd) do begin
    Fnd := (Pos(Lst.Strings[j], Str) > 0); Inc(j);
  end; Result := Fnd;
end;

// --------------------------------------------------------------------------
//
// --------------------------------------------------------------------------

function PadLeft(Str: String; Chr: Char; Len: Integer): String;
begin
  while (Length(Str) < Len) do Str := Chr + Str; Result := Str;
end;

// --------------------------------------------------------------------------
//
// --------------------------------------------------------------------------

function PadRight(Str: String; Chr: Char; Len: Integer): String;
begin
  while (Length(Str) < Len) do Str := Str + Chr; Result := Str;
end;

// --------------------------------------------------------------------------
//
// --------------------------------------------------------------------------

function IIFString(Condition: Boolean; TrueCase, FalseCase: String): String;
begin
  if Condition then Result := TrueCase else Result := FalseCase;
end;

// --------------------------------------------------------------------------
//
// --------------------------------------------------------------------------

function RepeatString(Str: String; Times: Integer): String;
var
  j: Integer;
  Tmp: String;
begin
  SetLength(Tmp, 0); for j := 1 to Times do Tmp := Tmp + Str; Result := Tmp;
end;

// --------------------------------------------------------------------------
//
// --------------------------------------------------------------------------

function CustomSortCompareName(Node1, Node2: TTreeNode; Param: Longint): Integer; stdcall;
begin
  if (not(PFileInfo(Node1.Data)^.Directory) and PFileInfo(Node2.Data)^.Directory) then begin Result := 1; exit; end;
  if (PFileInfo(Node1.Data)^.Directory and not(PFileInfo(Node2.Data)^.Directory)) then begin Result := -1; exit; end;
  if (LowerCase(Node1.Text) > LowerCase(Node2.Text)) then Result := 1 else if (LowerCase(Node1.Text) < LowerCase(Node2.Text)) then Result := -1 else Result := 0; if (Param = 1) then Result := -1 * Result;
end;

// --------------------------------------------------------------------------
//
// --------------------------------------------------------------------------

function CustomSortCompareSize(Node1, Node2: TTreeNode; Param: Longint): Integer; stdcall;
begin
  if (not(PFileInfo(Node1.Data)^.Directory) and PFileInfo(Node2.Data)^.Directory) then begin Result := 1; exit; end;
  if (PFileInfo(Node1.Data)^.Directory and not(PFileInfo(Node2.Data)^.Directory)) then begin Result := -1; exit; end;
  if (PFileInfo(Node1.Data)^.Size > PFileInfo(Node2.Data)^.Size) then Result := 1 else if (PFileInfo(Node1.Data)^.Size < PFileInfo(Node2.Data)^.Size) then Result := -1 else if (LowerCase(Node1.Text) > LowerCase(Node2.Text)) then Result := 1 else if (LowerCase(Node1.Text) < LowerCase(Node2.Text)) then Result := -1 else Result := 0; if (Param = 1) then Result := -1 * Result;
end;

// --------------------------------------------------------------------------
//
// --------------------------------------------------------------------------

function CustomSortCompareCreationTime(Node1, Node2: TTreeNode; Param: Longint): Integer; stdcall;
begin
  if (not(PFileInfo(Node1.Data)^.Directory) and PFileInfo(Node2.Data)^.Directory) then begin Result := 1; exit; end;
  if (PFileInfo(Node1.Data)^.Directory and not(PFileInfo(Node2.Data)^.Directory)) then begin Result := -1; exit; end;
  if (PFileInfo(Node1.Data)^.CreationTime > PFileInfo(Node2.Data)^.CreationTime) then Result := 1 else if (PFileInfo(Node1.Data)^.CreationTime < PFileInfo(Node2.Data)^.CreationTime) then Result := -1 else if (LowerCase(Node1.Text) > LowerCase(Node2.Text)) then Result := 1 else if (LowerCase(Node1.Text) < LowerCase(Node2.Text)) then Result := -1 else Result := 0; if (Param = 1) then Result := -1 * Result;
end;

// --------------------------------------------------------------------------
//
// --------------------------------------------------------------------------

function CustomSortCompareLastModifiedTime(Node1, Node2: TTreeNode; Param: Longint): Integer; stdcall;
begin
  if (not(PFileInfo(Node1.Data)^.Directory) and PFileInfo(Node2.Data)^.Directory) then begin Result := 1; exit; end;
  if (PFileInfo(Node1.Data)^.Directory and not(PFileInfo(Node2.Data)^.Directory)) then begin Result := -1; exit; end;
  if (PFileInfo(Node1.Data)^.LastModifiedTime > PFileInfo(Node2.Data)^.LastModifiedTime) then Result := 1 else if (PFileInfo(Node1.Data)^.LastModifiedTime < PFileInfo(Node2.Data)^.LastModifiedTime) then Result := -1 else if (LowerCase(Node1.Text) > LowerCase(Node2.Text)) then Result := 1 else if (LowerCase(Node1.Text) < LowerCase(Node2.Text)) then Result := -1 else Result := 0; if (Param = 1) then Result := -1 * Result;
end;

// --------------------------------------------------------------------------
//
// --------------------------------------------------------------------------

function CustomSortCompareLastAccessedTime(Node1, Node2: TTreeNode; Param: Longint): Integer; stdcall;
begin
  if (not(PFileInfo(Node1.Data)^.Directory) and PFileInfo(Node2.Data)^.Directory) then begin Result := 1; exit; end;
  if (PFileInfo(Node1.Data)^.Directory and not(PFileInfo(Node2.Data)^.Directory)) then begin Result := -1; exit; end;
  if (PFileInfo(Node1.Data)^.LastAccessedTime > PFileInfo(Node2.Data)^.LastAccessedTime) then Result := 1 else if (PFileInfo(Node1.Data)^.LastAccessedTime < PFileInfo(Node2.Data)^.LastAccessedTime) then Result := -1 else if (LowerCase(Node1.Text) > LowerCase(Node2.Text)) then Result := 1 else if (LowerCase(Node1.Text) < LowerCase(Node2.Text)) then Result := -1 else Result := 0; if (Param = 1) then Result := -1 * Result;
end;

// --------------------------------------------------------------------------
//
// --------------------------------------------------------------------------

function CustomSortCompareScramble(Node1, Node2: TTreeNode; Param: Longint): Integer; stdcall;
begin
  if (not(PFileInfo(Node1.Data)^.Directory) and PFileInfo(Node2.Data)^.Directory) then begin Result := 1; exit; end;
  if (PFileInfo(Node1.Data)^.Directory and not(PFileInfo(Node2.Data)^.Directory)) then begin Result := -1; exit; end;
  Result := Random(3) - 1;
end;

// --------------------------------------------------------------------------
//
// --------------------------------------------------------------------------

function GetCompletePathOfTreeNode(Node: TTreeNode): String;
var
  Path: String;
begin
  Path := ''; Node := Node.Parent; while (Node <> nil) do begin
    if (Length(Path) > 0) then Path := Node.Text + '\' + Path else Path := Node.Text; Node := Node.Parent;
  end; Result := Path;
end;

// --------------------------------------------------------------------------
//
// --------------------------------------------------------------------------

function HasNonDirectoryChildren(Node: TTreeNode): Boolean;
begin
  Result := False; if Node.HasChildren then begin
    Node := Node.GetFirstChild; while (Node <> nil) do begin
      if PFileInfo(Node.Data)^.Directory then begin
        if HasNonDirectoryChildren(Node) then begin Result := True; exit; end;
      end else begin
        Result := True; exit;
      end; Node := Node.Parent.GetNextChild(Node);
    end;
  end;
end;

// --------------------------------------------------------------------------
//
// --------------------------------------------------------------------------

procedure TfrmStartup.ApplyStyleGoodOldTree(Daddy: TTreeNode; Level: Integer);
var
  Node: TTreeNode;
begin
  if (Daddy <> nil) then Node := Daddy.GetFirstChild else Node := trTree.Items[0]; while (Node <> nil) do begin
    if not(PFileInfo(Node.Data)^.Directory) or not(pmHideEmptyFolders) or HasNonDirectoryChildren(Node) then lsOutput.Append(RepeatString('|   ', Level - 1) + '|-- ' + Node.Text + IIFString(PFileInfo(Node.Data)^.Directory, ' <DIR>', ''));
    if Node.HasChildren then ApplyStyleGoodOldTree(Node, Level + 1);
    Node := Daddy.GetNextChild(Node);
  end;
end;

// --------------------------------------------------------------------------
//
// --------------------------------------------------------------------------

procedure TfrmStartup.ApplyStyleHtmlTreeBriefPrintable(Daddy: TTreeNode; Level: Integer);
var
  Temp: String;
  Node: TTreeNode;
begin
  if (Daddy <> nil) then Node := Daddy.GetFirstChild else Node := trTree.Items[0]; while (Node <> nil) do begin
    if Node.HasChildren then begin Temp := GetCompletePathOfTreeNode(Node); if (Length(Temp) > 0) then Temp := '&nbsp;<i><small>(' + Temp + ')</small></i>'; end else Temp := '';
    if not(PFileInfo(Node.Data)^.Directory) or not(pmHideEmptyFolders) or HasNonDirectoryChildren(Node) then lsOutput.Append(RepeatString('|&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;', Level - 1) + IIFString(PFileInfo(Node.Data)^.Directory, '<b>' + Node.Text + '</b>', Node.Text) + Temp + '<br>');
    if Node.HasChildren then ApplyStyleHtmlTreeBriefPrintable(Node, Level + 1);
    Node := Daddy.GetNextChild(Node);
  end;
end;

// --------------------------------------------------------------------------
//
// --------------------------------------------------------------------------

procedure TfrmStartup.ApplyStyleHtmlTreeDetailedPrintable(Daddy: TTreeNode; Level: Integer);
var
  Temp: String;
  Node: TTreeNode;
begin
  if (Daddy <> nil) then Node := Daddy.GetFirstChild else Node := trTree.Items[0]; while (Node <> nil) do begin
    if Node.HasChildren then begin Temp := GetCompletePathOfTreeNode(Node); if (Length(Temp) > 0) then Temp := '&nbsp;<i><small>(' + Temp + ')</small></i>'; end else Temp := '';
    if not(PFileInfo(Node.Data)^.Directory) or not(pmHideEmptyFolders) or HasNonDirectoryChildren(Node) then lsOutput.Append('<tr><td>' + RepeatString('|&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;', Level - 1) + IIFString(PFileInfo(Node.Data)^.Directory, '<b>' + Node.Text + '</b>', Node.Text) + Temp + '</td><td>&nbsp;&nbsp;&nbsp;&nbsp;</td><td align=right>' + IIFString(not(PFileInfo(Node.Data)^.Directory), FormatCurr('###,##0', Round(0.0009765625 * PFileInfo(Node.Data)^.Size)) + ' KB', '') + '</td><td>&nbsp;&nbsp;&nbsp;&nbsp;</td><td>' + FormatDateTime('dd/mm/yyyy hh.nn', PFileInfo(Node.Data)^.LastModifiedTime) + '</td></tr>');
    if Node.HasChildren then ApplyStyleHtmlTreeDetailedPrintable(Node, Level + 1);
    Node := Daddy.GetNextChild(Node);
  end;
end;

// --------------------------------------------------------------------------
//
// --------------------------------------------------------------------------

procedure TfrmStartup.ApplyStyleHtmlTreeEverythingPrintable(Daddy: TTreeNode; Level: Integer);
var
  Temp: String;
  Node: TTreeNode;
begin
  if (Daddy <> nil) then Node := Daddy.GetFirstChild else Node := trTree.Items[0]; while (Node <> nil) do begin
    if Node.HasChildren then begin Temp := GetCompletePathOfTreeNode(Node); if (Length(Temp) > 0) then Temp := '&nbsp;<i><small>(' + Temp + ')</small></i>'; end else Temp := '';
    if not(PFileInfo(Node.Data)^.Directory) or not(pmHideEmptyFolders) or HasNonDirectoryChildren(Node) then lsOutput.Append('<tr><td>' + RepeatString('|&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;', Level - 1) + IIFString(PFileInfo(Node.Data)^.Directory, '<b>' + Node.Text + '</b>', Node.Text) + Temp + '</td><td>&nbsp;&nbsp;&nbsp;&nbsp;</td><td align=right>' + IIFString(not(PFileInfo(Node.Data)^.Directory), FormatCurr('###,##0', Round(0.0009765625 * PFileInfo(Node.Data)^.Size)) + ' KB', '') + '</td><td>&nbsp;&nbsp;&nbsp;&nbsp;</td><td>' + FormatDateTime('dd/mm/yyyy hh.nn', PFileInfo(Node.Data)^.LastModifiedTime) + '</td><td>&nbsp;&nbsp;&nbsp;&nbsp;</td><td>' + FormatDateTime('dd/mm/yyyy hh.nn', PFileInfo(Node.Data)^.LastAccessedTime) + '</td><td>&nbsp;&nbsp;&nbsp;&nbsp;</td><td>' + FormatDateTime('dd/mm/yyyy hh.nn', PFileInfo(Node.Data)^.CreationTime) + '</td></tr>');
    if Node.HasChildren then ApplyStyleHtmlTreeEverythingPrintable(Node, Level + 1);
    Node := Daddy.GetNextChild(Node);
  end;
end;

// --------------------------------------------------------------------------
//
// --------------------------------------------------------------------------

procedure TfrmStartup.ApplyStyleHtmlTreeDetailedMultimedia(Daddy: TTreeNode; Level: Integer);
var
  Node: TTreeNode;
  Path, Temp: String;
begin
  if (Daddy <> nil) then Node := Daddy.GetFirstChild else Node := trTree.Items[0]; while (Node <> nil) do begin
    Path := GetCompletePathOfTreeNode(Node);
    if Node.HasChildren then begin if (Length(Path) > 0) then Temp := '&nbsp;<i><small>(' + Path + ')</small></i>'; end else Temp := '';
    if not(PFileInfo(Node.Data)^.Directory) or not(pmHideEmptyFolders) or HasNonDirectoryChildren(Node) then begin lsOutput.Append('<tr bgcolor="' + IIFString((Self.Tag mod 2) = 0, '#f0f0ff', '#e0e0ff') + '"><td>&nbsp;&nbsp;&nbsp;&nbsp;' + RepeatString('|&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;', Level - 1) + IIFString(PFileInfo(Node.Data)^.Directory, '<b>' + Node.Text + '</b>', '<a href="' + pmFolder + '\' + Path + '\' + Node.Text + '">' + Node.Text + '</a>') + Temp + '</td><td>&nbsp;&nbsp;&nbsp;&nbsp;</td><td align=right>' + IIFString(not(PFileInfo(Node.Data)^.Directory), FormatCurr('###,##0', Round(0.0009765625 * PFileInfo(Node.Data)^.Size)) + ' KB', '') + '</td><td>&nbsp;&nbsp;&nbsp;&nbsp;</td><td>' + FormatDateTime('dd/mm/yyyy hh.nn', PFileInfo(Node.Data)^.LastModifiedTime) + '&nbsp;&nbsp;&nbsp;&nbsp;</td></tr>'); Self.Tag := Self.Tag + 1; end;
    if Node.HasChildren then ApplyStyleHtmlTreeDetailedMultimedia(Node, Level + 1);
    Node := Daddy.GetNextChild(Node);
  end;
end;

// --------------------------------------------------------------------------
//
// --------------------------------------------------------------------------

procedure TfrmStartup.ApplyStyleHtmlTreeCompactMultimedia(Daddy: TTreeNode; Level: Integer);
var
  Node: TTreeNode;
  Path, Temp: String;
begin
  if (Daddy <> nil) then Node := Daddy.GetFirstChild else Node := trTree.Items[0]; while (Node <> nil) do begin
    Path := GetCompletePathOfTreeNode(Node);
    if Node.HasChildren then begin if (Length(Path) > 0) then Temp := '&nbsp;<i><small>(' + Path + ')</small></i>'; end else Temp := '';
    if not(PFileInfo(Node.Data)^.Directory) then lsOutput.Append('<a href="' + pmFolder + '\' + Path + '\' + Node.Text + '">' + Node.Text + '</a>;') else if not(pmHideEmptyFolders) or HasNonDirectoryChildren(Node) then begin lsOutput.Append('</td></tr></table><table width=100% cellspacing=0 cellpadding=0><tr bgcolor="' + IIFString((Self.Tag mod 2) = 0, '#f0f0ff', '#e0e0ff') + '"><td width=' + IntToStr(16 * (Level - 1)) + '></td><td><b>' + Node.Text + '</b>' + Temp + '</td></tr></table><table width=100% cellspacing=0 cellpadding=0><tr bgcolor="' + IIFString((Self.Tag mod 2) = 0, '#f0f0ff', '#e0e0ff') + '"><td width=' + IntToStr(16 * Level) + '></td><td>'); Self.Tag := Self.Tag + 1; end;
    if Node.HasChildren then ApplyStyleHtmlTreeCompactMultimedia(Node, Level + 1);
    Node := Daddy.GetNextChild(Node);
  end;
end;

// --------------------------------------------------------------------------
//
// --------------------------------------------------------------------------

procedure TfrmStartup.ApplyStyleWindowsMediaPlayerPlaylist(Daddy: TTreeNode);
var
  Path: String;
  Node: TTreeNode;
begin
  if (Daddy <> nil) then Node := Daddy.GetFirstChild else Node := trTree.Items[0]; while (Node <> nil) do begin
    Path := GetCompletePathOfTreeNode(Node);
    if not(PFileInfo(Node.Data)^.Directory) then lsOutput.Append('<Entry><Ref href="' + pmFolder + '\' + Path + '\' + Node.Text + '" /></Entry>');
    if Node.HasChildren then ApplyStyleWindowsMediaPlayerPlaylist(Node);
    Node := Daddy.GetNextChild(Node);
  end;
end;

// --------------------------------------------------------------------------
//
// --------------------------------------------------------------------------

procedure TfrmStartup.ApplyStyleGeneric;
begin
  case pmStyle of
    0: begin lsOutput.Append('Contents of ' + pmFolder); lsOutput.Append(''); ApplyStyleGoodOldTree(nil, 1); pmDestination := 'Results.txt'; end;
    1: begin lsOutput.Append('<html><style>body,td{font-family:Tahoma;font-size:11px}small{font-size:9px}</style><body bgcolor="#ffffff">Contents of <b>' + pmFolder + '</b><br><br>'); ApplyStyleHtmlTreeBriefPrintable(nil, 1); lsOutput.Append('</body></html>'); pmDestination := 'Results.htm'; end;
    2: begin lsOutput.Append('<html><style>body,td{font-family:Tahoma;font-size:11px}small{font-size:9px}</style><body bgcolor="#ffffff">Contents of <b>' + pmFolder + '</b><br><table cellspacing=0 cellpadding=0><tr><td></td><td></td><td align=right><b>Size</b></td><td></td><td><b>Date Modified</b></td></tr>'); ApplyStyleHtmlTreeDetailedPrintable(nil, 1); lsOutput.Append('</table></body></html>'); pmDestination := 'Results.htm'; end;
    3: begin lsOutput.Append('<html><style>body,td{font-family:Tahoma;font-size:11px}small{font-size:9px}</style><body bgcolor="#ffffff">Contents of <b>' + pmFolder + '</b><br><table cellspacing=0 cellpadding=0><tr><td></td><td></td><td align=right><b>Size</b></td><td></td><td><b>Date Modified</b></td><td></td><td><b>Date Accessed</b></td><td></td><td><b>Date Created</b></td></tr>'); ApplyStyleHtmlTreeEverythingPrintable(nil, 1); lsOutput.Append('</table></body></html>'); pmDestination := 'Results.htm'; end;
    4: begin lsOutput.Append('<html><style>body,td{font-family:Tahoma;font-size:11px}small{font-size:9px}</style><body bgcolor="#ffffff" link="#000080" vlink="#000080">Contents of <b>' + pmFolder + '</b><br><table width=100% cellspacing=0 cellpadding=0><tr><td></td><td></td><td align=right><b>Size</b></td><td></td><td><b>Date Modified</b></td></tr>'); ApplyStyleHtmlTreeDetailedMultimedia(nil, 1); lsOutput.Append('</table></body></html>'); pmDestination := 'Results.htm'; end;
    5: begin lsOutput.Append('<html><style>body,td{font-family:Tahoma;font-size:11px}small{font-size:9px}</style><body bgcolor="#ffffff" link="#000080" vlink="#000080">Contents of <b>' + pmFolder + '</b><br><br><table width=100% cellspacing=0 cellpadding=0><tr><td>'); ApplyStyleHtmlTreeCompactMultimedia(nil, 1); lsOutput.Append('</body></html>'); pmDestination := 'Results.htm'; end;
    6: begin lsOutput.Append('<ASX version="3.0">'); ApplyStyleWindowsMediaPlayerPlaylist(nil); lsOutput.Append('</ASX>'); pmDestination := 'Results.asx'; end;
  end;
end;

// --------------------------------------------------------------------------
//
// --------------------------------------------------------------------------

function FileTimeToDateTime(Value: TFileTime): TDateTime;
var
  Tmp: TSystemTime;
begin
  FileTimeToSystemTime(Value, Tmp); Result := SystemTimeToDateTime(Tmp);
end;

// --------------------------------------------------------------------------
//
// --------------------------------------------------------------------------

function TfrmStartup.ApplyFilters(Path: String; Data: TSearchRec): Boolean;
begin
  Result := True;
  if (Result) and (pmHideFiles) and ((Data.Attr and faDirectory) = 0) then Result := False;
  if (Result) and (pmNameContains.Count > 0) and not(StringContainsListSubstr(Data.Name, pmNameContains)) then Result := False;
  if (Result) and (pmNameDoesNotContain.Count > 0) and StringContainsListSubstr(Data.Name, pmNameDoesNotContain) then Result := False;
  if (Result) and (pmSmallerThan > -1) and (Data.Size >= pmSmallerThan) then Result := False;
  if (Result) and (pmGreaterThan > -1) and (Data.Size <= pmGreaterThan) then Result := False;
  if (Result) and (pmEnabledCreated) and (FileTimeToDateTime(Data.FindData.ftCreationTime) >= pmCreatedBefore) then Result := False;
  if (Result) and (pmEnabledCreated) and (FileTimeToDateTime(Data.FindData.ftCreationTime) <= pmCreatedAfter) then Result := False;
  if (Result) and (pmEnabledLastModified) and (FileTimeToDateTime(Data.FindData.ftLastWriteTime) >= pmLastModifiedBefore) then Result := False;
  if (Result) and (pmEnabledLastModified) and (FileTimeToDateTime(Data.FindData.ftLastWriteTime) <= pmLastModifiedAfter) then Result := False;
  if (Result) and (pmEnabledLastAccessed) and (FileTimeToDateTime(Data.FindData.ftLastAccessTime) >= pmLastAccessedBefore) then Result := False;
  if (Result) and (pmEnabledLastAccessed) and (FileTimeToDateTime(Data.FindData.ftLastAccessTime) <= pmLastAccessedAfter) then Result := False;
end;

// --------------------------------------------------------------------------
//
// --------------------------------------------------------------------------

procedure TfrmStartup.SearchScan(Path: String; Daddy: TTreeNode);
var
  j: Integer;
  Refp: PFileInfo;
  Data: TSearchRec;
  Node: TTreeNode;
begin
  j := FindFirst(Path + '\*.*', faAnyFile, Data); while (j = 0) do begin
    if (Data.Name <> '.') and (Data.Name <> '..') then begin
      if ((Data.Attr and faDirectory) > 0) or ApplyFilters(Path, Data) then begin

        // Reference pointer
        GetMem(Refp, SizeOf(TFileInfo));
        Refp^.Size := Data.Size; Refp^.Directory := (Data.Attr and faDirectory > 0);
        Refp^.CreationTime := FileTimeToDateTime(Data.FindData.ftCreationTime);
        Refp^.LastModifiedTime := FileTimeToDateTime(Data.FindData.ftLastWriteTime);
        Refp^.LastAccessedTime := FileTimeToDateTime(Data.FindData.ftLastAccessTime);

        // Adding the pointer to the node list
        Node := trTree.Items.AddChild(Daddy, Data.Name); Node.Data := Refp;

        // Apply the search scan recursively
        if ((Data.Attr and faDirectory) > 0) then SearchScan(Path + '\' + Data.Name, Node);
        
      end;
    end; j := FindNext(Data);
  end; FindClose(Data);
end;

// --------------------------------------------------------------------------
//
// --------------------------------------------------------------------------

procedure TfrmStartup.SearchClearNodeData(Daddy: TTreeNode);
var
  Node: TTreeNode;
begin
  if (Daddy <> nil) then Node := Daddy.GetFirstChild else Node := trTree.Items[0]; while (Node <> nil) do begin
    FreeMem(Node.Data, SizeOf(TFileInfo));
    if Node.HasChildren then SearchClearNodeData(Node);
    Node := Daddy.GetNextChild(Node);
  end;
end;

// --------------------------------------------------------------------------
//
// --------------------------------------------------------------------------

procedure TfrmStartup.SearchClear;
begin
  // Freeing memory used by search structures
  if (trTree.Items.Count > 0) then SearchClearNodeData(nil); trTree.Items.Clear; lsOutput.Clear;
end;

// --------------------------------------------------------------------------
//
// --------------------------------------------------------------------------

procedure TfrmStartup.LowerCaseTree(Daddy: TTreeNode);
var
  Node: TTreeNode;
begin
  if (Daddy <> nil) then Node := Daddy.GetFirstChild else Node := trTree.Items[0]; while (Node <> nil) do begin
    Node.Text := LowerCase(Node.Text);
    if Node.HasChildren then LowerCaseTree(Node);
    Node := Daddy.GetNextChild(Node);
  end;
end;

// --------------------------------------------------------------------------
//
// --------------------------------------------------------------------------

procedure TfrmStartup.UpperCaseTree(Daddy: TTreeNode);
var
  Node: TTreeNode;
begin
  if (Daddy <> nil) then Node := Daddy.GetFirstChild else Node := trTree.Items[0]; while (Node <> nil) do begin
    Node.Text := UpperCase(Node.Text);
    if Node.HasChildren then UpperCaseTree(Node);
    Node := Daddy.GetNextChild(Node);
  end;
end;

// --------------------------------------------------------------------------
//
// --------------------------------------------------------------------------

procedure TfrmStartup.ProperCaseTree(Daddy: TTreeNode);
var
  Node: TTreeNode;
begin
  if (Daddy <> nil) then Node := Daddy.GetFirstChild else Node := trTree.Items[0]; while (Node <> nil) do begin
    Node.Text := ProperCase(Node.Text);
    if Node.HasChildren then ProperCaseTree(Node);
    Node := Daddy.GetNextChild(Node);
  end;
end;

// --------------------------------------------------------------------------
//
// --------------------------------------------------------------------------

procedure TfrmStartup.SearchPostProcess;
begin
  if pmLowerCase then LowerCaseTree(nil);
  if pmUpperCase then UpperCaseTree(nil);
  if pmProperCase then ProperCaseTree(nil);
end;

// --------------------------------------------------------------------------
//
// --------------------------------------------------------------------------

procedure TfrmStartup.SearchShowResults;
begin
  // Saving search results
  lsOutput.SaveToFile(AppPath + pmDestination);

  // Calling appropriate viewer
  ShellExecute(Handle, 'open', PChar(AppPath + pmDestination), '', '', SW_NORMAL);
end;

// --------------------------------------------------------------------------
//
// --------------------------------------------------------------------------

procedure TfrmStartup.SearchSort;
begin
  if (trTree.Items.Count > 1) then begin
    if pmSortingNone then exit;
    if pmSortingByName then trTree.CustomSort(@CustomSortCompareName, Integer(pmSortingReverseOrder));
    if pmSortingBySize then trTree.CustomSort(@CustomSortCompareSize, Integer(pmSortingReverseOrder));
    if pmSortingByCreationTime then trTree.CustomSort(@CustomSortCompareCreationTime, Integer(pmSortingReverseOrder));
    if pmSortingByLastModifiedTime then trTree.CustomSort(@CustomSortCompareLastModifiedTime, Integer(pmSortingReverseOrder));
    if pmSortingByLastAccessedTime then trTree.CustomSort(@CustomSortCompareLastAccessedTime, Integer(pmSortingReverseOrder));
    if pmSortingScramble then trTree.CustomSort(@CustomSortCompareScramble, 0);
  end;
end;

// --------------------------------------------------------------------------
//
// --------------------------------------------------------------------------

procedure TfrmStartup.btnBrowseClick(Sender: TObject);
begin
  // Gathering options
  pmFolder := edFolder.Text;
  pmStyle := edStyle.ItemIndex;
  Split(edNameContains.Text, ';', pmNameContains);
  Split(edNameDoesNotContain.Text, ';', pmNameDoesNotContain);
  pmSmallerThan := StrToIntDef(edSmallerThan.Text, -1);
  pmGreaterThan := StrToIntDef(edGreaterThan.Text, -1);
  pmCreatedBefore := edCreatedBefore.Date;
  pmCreatedAfter := edCreatedAfter.Date;
  pmLastModifiedBefore := edLastModifiedBefore.Date;
  pmLastModifiedAfter := edLastModifiedAfter.Date;
  pmLastAccessedBefore := edLastAccessedBefore.Date;
  pmLastAccessedAfter := edLastAccessedAfter.Date;
  pmEnabledCreated := cbEnabledCreated.Checked;
  pmEnabledLastModified := cbEnabledLastModified.Checked;
  pmEnabledLastAccessed := cbEnabledLastAccessed.Checked;
  pmHideEmptyFolders := cbHideEmptyFolders.Checked;
  pmHideFiles := cbHideFiles.Checked;
  pmSortingNone := rbSortingNone.Checked;
  pmSortingByName := rbSortingByName.Checked;
  pmSortingBySize := rbSortingBySize.Checked;
  pmSortingByCreationTime := rbSortingByCreationTime.Checked;
  pmSortingByLastModifiedTime := rbSortingByLastModifiedTime.Checked;
  pmSortingByLastAccessedTime := rbSortingByLastAccessedTime.Checked;
  pmSortingScramble := rbSortingScramble.Checked;
  pmSortingReverseOrder := cbSortingReverseOrder.Checked;
  pmLowerCase := rbLowerCase.Checked;
  pmUpperCase := rbUpperCase.Checked;
  pmProperCase := rbProperCase.Checked;
  pmUnchangedCase := rbUnchangedCase.Checked;

  // Updating status bar
  sbInfo.SimpleText := 'Operation in progress. Please wait...'; Application.ProcessMessages;

  // Search in progress...
  SearchClear; SearchScan(pmFolder, nil); SearchSort; SearchPostProcess;

  // Finalizing search results
  ApplyStyleGeneric; SearchShowResults; SearchClear;

  // Updating status bar
  sbInfo.SimpleText := ''; Application.ProcessMessages;
end;

// --------------------------------------------------------------------------
//
// --------------------------------------------------------------------------

procedure TfrmStartup.FormCreate(Sender: TObject);
begin
  // Initializing variables and random seed
  AppPath := ExtractFilePath(ParamStr(0)); Randomize; edStyle.ItemIndex := 0;
  trTree := TTreeView.Create(frmStartup); trTree.Visible := False; trTree.Parent := frmStartup;
  pmNameContains := TStringList.Create; pmNameDoesNotContain := TStringList.Create; lsOutput := TStringList.Create;
end;

// --------------------------------------------------------------------------
//
// --------------------------------------------------------------------------

procedure TfrmStartup.btnClearNameContainsClick(Sender: TObject);
begin
  edNameContains.Text := '';
end;

// --------------------------------------------------------------------------
//
// --------------------------------------------------------------------------

procedure TfrmStartup.btnClearNameDoesNotContainClick(Sender: TObject);
begin
  edNameDoesNotContain.Text := '';
end;

// --------------------------------------------------------------------------
//
// --------------------------------------------------------------------------

procedure TfrmStartup.LoadProfile(FileName: String);
var
  Params: TStringList;
begin
  Params := TStringList.Create;
  Params.LoadFromFile(FileName);
  if (Params.Count >= 25) then begin
    edFolder.Text := Params.Strings[0];
    edStyle.ItemIndex := StrToIntDef(Params.Strings[1], 0);
    edNameContains.Text := Params.Strings[2];
    edNameDoesNotContain.Text := Params.Strings[3];
    edSmallerThan.Text := Params.Strings[4];
    edGreaterThan.Text := Params.Strings[5];
    edCreatedBefore.Date := StrToDate(Params.Strings[6]);
    edCreatedAfter.Date := StrToDate(Params.Strings[7]);
    edLastModifiedBefore.Date := StrToDate(Params.Strings[8]);
    edLastModifiedAfter.Date := StrToDate(Params.Strings[9]);
    edLastAccessedBefore.Date := StrToDate(Params.Strings[10]);
    edLastAccessedAfter.Date := StrToDate(Params.Strings[11]);
    cbHideEmptyFolders.Checked := Boolean(StrToIntDef(Params.Strings[12], 0));
    rbSortingNone.Checked := Boolean(StrToIntDef(Params.Strings[13], 0));
    rbSortingByName.Checked := Boolean(StrToIntDef(Params.Strings[14], 0));
    rbSortingBySize.Checked := Boolean(StrToIntDef(Params.Strings[15], 0));
    rbSortingByCreationTime.Checked := Boolean(StrToIntDef(Params.Strings[16], 0));
    rbSortingByLastModifiedTime.Checked := Boolean(StrToIntDef(Params.Strings[17], 0));
    rbSortingByLastAccessedTime.Checked := Boolean(StrToIntDef(Params.Strings[18], 0));
    rbSortingScramble.Checked := Boolean(StrToIntDef(Params.Strings[19], 0));
    cbSortingReverseOrder.Checked := Boolean(StrToIntDef(Params.Strings[20], 0));
    rbLowerCase.Checked := Boolean(StrToIntDef(Params.Strings[21], 0));
    rbUpperCase.Checked := Boolean(StrToIntDef(Params.Strings[22], 0));
    rbProperCase.Checked := Boolean(StrToIntDef(Params.Strings[23], 0));
    rbUnchangedCase.Checked := Boolean(StrToIntDef(Params.Strings[24], 0));
    if (Params.Count >= 26) then cbHideFiles.Checked := Boolean(StrToIntDef(Params.Strings[25], 0));
    if (Params.Count >= 27) then cbEnabledCreated.Checked := Boolean(StrToIntDef(Params.Strings[26], 0));
    if (Params.Count >= 28) then cbEnabledLastModified.Checked := Boolean(StrToIntDef(Params.Strings[27], 0));
    if (Params.Count >= 29) then cbEnabledLastAccessed.Checked := Boolean(StrToIntDef(Params.Strings[28], 0));
  end;
  Params.Free;
end;

// --------------------------------------------------------------------------
//
// --------------------------------------------------------------------------

procedure TfrmStartup.SaveProfileAs(FileName: String);
var
  Params: TStringList;
begin
  Params := TStringList.Create;
  Params.Append(edFolder.Text);
  Params.Append(IntToStr(edStyle.ItemIndex));
  Params.Append(edNameContains.Text);
  Params.Append(edNameDoesNotContain.Text);
  Params.Append(edSmallerThan.Text);
  Params.Append(edGreaterThan.Text);
  Params.Append(DateToStr(edCreatedBefore.Date));
  Params.Append(DateToStr(edCreatedAfter.Date));
  Params.Append(DateToStr(edLastModifiedBefore.Date));
  Params.Append(DateToStr(edLastModifiedAfter.Date));
  Params.Append(DateToStr(edLastAccessedBefore.Date));
  Params.Append(DateToStr(edLastAccessedAfter.Date));
  Params.Append(IntToStr(Integer(cbHideEmptyFolders.Checked)));
  Params.Append(IntToStr(Integer(rbSortingNone.Checked)));
  Params.Append(IntToStr(Integer(rbSortingByName.Checked)));
  Params.Append(IntToStr(Integer(rbSortingBySize.Checked)));
  Params.Append(IntToStr(Integer(rbSortingByCreationTime.Checked)));
  Params.Append(IntToStr(Integer(rbSortingByLastModifiedTime.Checked)));
  Params.Append(IntToStr(Integer(rbSortingByLastAccessedTime.Checked)));
  Params.Append(IntToStr(Integer(rbSortingScramble.Checked)));
  Params.Append(IntToStr(Integer(cbSortingReverseOrder.Checked)));
  Params.Append(IntToStr(Integer(rbLowerCase.Checked)));
  Params.Append(IntToStr(Integer(rbUpperCase.Checked)));
  Params.Append(IntToStr(Integer(rbProperCase.Checked)));
  Params.Append(IntToStr(Integer(rbUnchangedCase.Checked)));
  Params.Append(IntToStr(Integer(cbHideFiles.Checked)));
  Params.Append(IntToStr(Integer(cbEnabledCreated.Checked)));
  Params.Append(IntToStr(Integer(cbEnabledLastModified.Checked)));
  Params.Append(IntToStr(Integer(cbEnabledLastAccessed.Checked)));
  Params.SaveToFile(FileName);
  Params.Free;
end;

// --------------------------------------------------------------------------
//
// --------------------------------------------------------------------------

procedure TfrmStartup.btnLoadProfileClick(Sender: TObject);
begin
  dlgLoadProfile.InitialDir := AppPath; if dlgLoadProfile.Execute then LoadProfile(dlgLoadProfile.FileName);
end;

// --------------------------------------------------------------------------
//
// --------------------------------------------------------------------------

procedure TfrmStartup.btnSaveProfileAsClick(Sender: TObject);
begin
  dlgSaveProfileAs.InitialDir := AppPath; if dlgSaveProfileAs.Execute then SaveProfileAs(dlgSaveProfileAs.FileName);
end;

// --------------------------------------------------------------------------
//
// --------------------------------------------------------------------------

procedure TfrmStartup.edNameContainsPrototypesChange(Sender: TObject);
begin
  if (Length(edNameContains.Text) > 0) then edNameContains.Text := edNameContains.Text + ';'; case edNameContainsPrototypes.ItemIndex of
    0: edNameContains.Text := edNameContains.Text + '.exe';
    1: edNameContains.Text := edNameContains.Text + '.zip;.arj;.rar';
    2: edNameContains.Text := edNameContains.Text + '.wav;.mp2;.mp3;.wma';
    3: edNameContains.Text := edNameContains.Text + '.bmp;.jpg;.gif;.png';
    4: edNameContains.Text := edNameContains.Text + '.txt;.doc;.xls';
    5: edNameContains.Text := edNameContains.Text + '.avi;.mpg;.wmv;.asf';
  end;
end;

// --------------------------------------------------------------------------
//
// --------------------------------------------------------------------------

procedure TfrmStartup.edNameDoesNotContainPrototypesChange(
  Sender: TObject);
begin
  if (Length(edNameDoesNotContain.Text) > 0) then edNameDoesNotContain.Text := edNameDoesNotContain.Text + ';'; case edNameDoesNotContainPrototypes.ItemIndex of
    0: edNameDoesNotContain.Text := edNameDoesNotContain.Text + '.exe';
    1: edNameDoesNotContain.Text := edNameDoesNotContain.Text + '.zip;.arj;.rar';
    2: edNameDoesNotContain.Text := edNameDoesNotContain.Text + '.wav;.mp2;.mp3;.wma';
    3: edNameDoesNotContain.Text := edNameDoesNotContain.Text + '.bmp;.jpg;.gif;.png';
    4: edNameDoesNotContain.Text := edNameDoesNotContain.Text + '.txt;.doc;.xls';
    5: edNameDoesNotContain.Text := edNameDoesNotContain.Text + '.avi;.mpg;.wmv;.asf';
  end;
end;

// --------------------------------------------------------------------------
//
// --------------------------------------------------------------------------

procedure TfrmStartup.btnChooseClick(Sender: TObject);
var
  Instance: TBrowseForFolder;
begin
  Instance := TBrowseForFolder.Create; Instance.Setup(edFolder.Text);
  if Instance.Execute('Select a folder', Handle) then edFolder.Text := Instance.GetFolder;
  Instance.Free;
end;

// --------------------------------------------------------------------------
//
// --------------------------------------------------------------------------

end.
