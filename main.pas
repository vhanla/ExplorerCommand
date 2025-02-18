unit main;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, SynEdit, DosCommand, rkShellPath, rkEdit,
  Vcl.StdCtrls, Vcl.ExtCtrls, System.ImageList, Vcl.ImgList,
  Vcl.ComCtrls, Vcl.WinXCtrls, TlHelp32, ShellApi, ShDocVw, ActiveX, ShlObj, IniFiles, ComObj,
  Vcl.Menus, DzDirSeek, rkSmartPath, rkVistaProBar, Vcl.VirtualImage,
  uHostPreview, Winapi.Wincodec, StrUtils, ES.BaseControls, ES.Images, rkView,
  JPEG, Math, CommCtrl {HIMAGELIST}, rkIntegerList, SynEditHighlighter,
  SynHighlighterUNIXShellScript, CB.Form, madExceptVcl, scStyledForm, libgit2,
  rkPathViewer, IconFontsImageListBase, IconFontsImageList, Clipbrd,
  SynHighlighterMulti, SynEditCodeFolding, SynHighlighterPas, Vcl.Buttons,
  System.Actions, Vcl.ActnList, Vcl.ToolWin, MPCommonObjects,
  EasyListview, VirtualExplorerEasyListview,
  Process, CB.Autorun, System.SyncObjs, ACL.UI.Controls.Base,
  ACL.UI.Controls.Labels, ACL.UI.Controls.ActivityIndicator,
  Vcl.VirtualImageList, Vcl.BaseImageCollection, Vcl.ImageCollection,
  ACL.UI.Controls.CompoundControl, ACL.UI.Controls.HexView, ACL.Classes,
  ACL.UI.Application;

const
  KeyEvent = WM_USER + 1;
  KeyEventAll = WM_USER + 2;
  CM_UpdateView = WM_USER + 2;
  CM_Progress   = WM_USER + 3;
  IID_IImageList: TGUID = '{46EB5926-582E-4017-9FDF-E8998DAA0950}';

type
  EInvalidImageFormat = class(Exception);

type

  PItemData = ^TItemData;
  TItemData = record
    Name: string;
    ThumbWidth: Word;
    ThumbHeight: Word;
    Size: Integer;
    Modified: TDateTime;
    Dir: Boolean;
    GotThumb: Boolean;
    IWidth, IHeight: Word;
    ImgIdx: Integer;
    IsIcon: Boolean;
    ImgState: Byte;
    Image: TObject;
  end;

  ThumbThread = class(TThread)
  private
    { Private Declarations }
    ViewLink: TrkView;
    ItemsLink: TList;
  protected
    procedure Execute; override;
  public
    constructor Create(View: TrkView; Items: TList);
  end;

  TFuzzyStringMatcher = class
  private
    FThreshold: Integer;
    function DamerauLevenshteinDistance(const S1, S2: string): Integer;
  public
    constructor Create(Threshold: Integer);
    function IsMatch(const Str, SubStr: string): Boolean;
  end;

  // Autocomplete https://stackoverflow.com/a/5465826
  TEnumString = class(TInterfacedObject, IEnumString)
  private
    type
      TPointerList = array[0..0] of Pointer;
    var
    FStrings: TStringList;
    FCurrIndex: Integer;
  public
    // IEnumString
    function Next(celt: Longint; out elt;
      pceltFetched: PLongint): HResult; stdcall;
    function Skip(celt: Longint): HResult; stdcall;
    function Reset: HResult; stdcall;
    function Clone(out enm: IEnumString): HResult; stdcall;
    // VCL
    constructor Create;
    destructor Destroy; override;
  end;

  {  ACO_NONE               = 0;
  ACO_AUTOSUGGEST        = $1;
  ACO_AUTOAPPEND         = $2;
  ACO_SEARCH             = $4;
  ACO_FILTERPREFIXES     = $8;
  ACO_USETAB             = $10;
  ACO_UPDOWNKEYDROPSLIST = $20;
  ACO_RTLREADING         = $40;
  ACO_WORD_FILTER        = $80;
  ACO_NOPREFIXFILTERING  = $100;
  }
  TACOption = (acNone, acAutoSuggest, acAutoAppend, acSearch, acFilterPrefixes,
               acUseTab, acUpDownKeyDropsList, acRTLReading, acWordFilter, acNoPrefixFiltering);
  TACOptions = set of TACOption;
  TACSource = (acsList, acsHistory, acsMRU, acsShell);
  TButtonedEdit = class(Vcl.ExtCtrls.TButtonedEdit)
  private
    FACList: TEnumString;
    FEnumString: IEnumString;
    FAutoComplete: IAutoComplete;
    FACEnabled: Boolean;
    FACOptions: TACOptions;
    FACSource: TACSource;
    function GetACStrings : TStringList;
    procedure SetACEnabled(const Value: Boolean);
    procedure SetACOptions(const Value: TACOptions);
    procedure SetACSource(const Value: TACSource);
    procedure SetACStrings(const Value: TStringList);
    class constructor Create;
  protected
    procedure CreateWnd; override;
    procedure DestroyWnd; override;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
  published
    property ACEnabled: Boolean read FACEnabled write SetACEnabled;
    property ACOptions: TACOptions read FACOptions write SetACOptions;
    property ACSource: TACSource read FACSource write SetACSource;
    property ACStrings: TStringList read GetACStrings write SetACStrings;
  end;

  TCommandType = (ctNormal, ctEnvironment, ctOther);

  TForm1 = class(TForm)
    DosCommand1: TDosCommand;
    ButtonedEdit1: TButtonedEdit;
    ImageList1: TImageList;
    BCEditor1: TSynEdit;
    StatusBar1: TStatusBar;
    SearchBox1: TSearchBox;
    TrayIcon1: TTrayIcon;
    PopupMenu1: TPopupMenu;
    Exit1: TMenuItem;
    Show1: TMenuItem;
    N1: TMenuItem;
    DzDirSeek1: TDzDirSeek;
    pnlPreview: TPanel;
    Splitter1: TSplitter;
    EsImage1: TEsImage;
    rkView1: TrkView;
    Image1: TImage;
    SynUNIXShellScriptSyn1: TSynUNIXShellScriptSyn;
    ListBox1: TListBox;
    ComboBox1: TComboBox;
    pnlTop: TPanel;
    IconFontsImageList1: TIconFontsImageList;
    rkSmartPath1: TrkSmartPath;
    PopupMenu2: TPopupMenu;
    OpenURL1: TMenuItem;
    CopyPathtoClipboard1: TMenuItem;
    SynPasSyn1: TSynPasSyn;
    SynMultiSyn1: TSynMultiSyn;
    SpeedButton1: TSpeedButton;
    IconFontsImageList2: TIconFontsImageList;
    ActionList1: TActionList;
    actPreview: TAction;
    actHide: TAction;
    ToolBar1: TToolBar;
    ToolButton1: TToolButton;
    Panel1: TPanel;
    actUnPin: TAction;
    actSigInt: TAction;
    VirtualMultiPathExplorerEasyListview1: TVirtualMultiPathExplorerEasyListview;
    actPath2Clip: TAction;
    tmrToast: TTimer;
    AppAutoStart1: TCBAutoStart;
    mnuAutoStart: TMenuItem;
    pnlTitle: TPanel;
    LinkLabel1: TLinkLabel;
    tmrOutput: TTimer;
    ActivityIndicator1: TActivityIndicator;
    ImageCollection1: TImageCollection;
    VirtualImageList1: TVirtualImageList;
    ACLHexView1: TACLHexView;
    ACLApplicationController1: TACLApplicationController;
    btnFileHandler: TSpeedButton;
    procedure ButtonedEdit1Enter(Sender: TObject);
    procedure ButtonedEdit1KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure DosCommand1ExecuteError(ASender: TObject; AE: Exception;
      var AHandled: Boolean);
    procedure DosCommand1NewLine(ASender: TObject; const ANewLine: string;
      AOutputType: TOutputType);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure Show1Click(Sender: TObject);
    procedure Exit1Click(Sender: TObject);
    procedure TrayIcon1DblClick(Sender: TObject);
    procedure DosCommand1Terminated(Sender: TObject);
    procedure DosCommand1TerminateProcess(ASender: TObject;
      var ACanTerminate: Boolean);
    procedure ListBox1DblClick(Sender: TObject);
    procedure ListBox1KeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure OpenURL1Click(Sender: TObject);
    procedure CopyPathtoClipboard1Click(Sender: TObject);
    procedure ButtonedEdit1KeyPress(Sender: TObject; var Key: Char);
    procedure BCEditor1DblClick(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure actPreviewExecute(Sender: TObject);
    procedure actUnPinExecute(Sender: TObject);
    procedure actSigIntExecute(Sender: TObject);
    procedure actPath2ClipExecute(Sender: TObject);
    procedure tmrToastTimer(Sender: TObject);
    procedure mnuAutoStartClick(Sender: TObject);
    procedure LinkLabel1LinkClick(Sender: TObject; const Link: string;
      LinkType: TSysLinkType);
    procedure tmrOutputTimer(Sender: TObject);
    procedure btnFileHandlerClick(Sender: TObject);
  private
    { Private declarations }
    FOutputBuffer: TStringList;
    FSyncLock: TCriticalSection;

    FPinned: Boolean;
    Items: TList;
    ThumbSizeW, ThumbSizeH: Integer;
    FhImageList48: Cardinal;
    FIconSize: Integer;

    FCommandOutput: TStringList;

    lastExplorerHandle: HWND;
    lastExplorerPath: String;
    lstExplorerPath: TStringList;
    lstExplorerWnd: TStringList;
    lstExplorerItem: TStringList;

    fPreview: THostPreviewHandler;
    fHexBuffer: TFileStream;

    function ListExplorerInstances:Integer;
    procedure KeyEventHandler(var Msg: TMessage); message KeyEvent;
    procedure KeyEventHandlerAll(var Msg: TMessage); message KeyEventAll;
    procedure OnFocusLost(Sender: TObject);

    function GetExplorerAddressBarRect(AHandle: HWND): TRect;
    function ShowPreview(const FileName: string): Boolean;
    procedure SwitchToWindow(AWnd: HWND);

    procedure ProcessDosCommand(Sender: TObject; ACommand: string; terminateCurrent: Boolean = False);

    procedure CMFocusChanged(var Msg: TCMFocusChanged); message CM_FOCUSCHANGED;

    procedure UpdateMainMenu(const ForeGroundWindow: HWND);

    procedure FlushIcons;

    procedure NoBorder(var Msg: TWMNCActivate); message WM_NCACTIVATE;
  protected
    procedure CreateParams(var Params: TCreateParams); override;
    procedure WndProc(var Message: TMessage); override;
  private
    FCommandType: TCommandType;
    FEnvExecutables: TStringList;
    FEnvStrings: TStringList;
    procedure UpdateStyle;

    procedure RefreshEnvironmentVariables;
    procedure WMSettingChange(var Msg: TMessage); message WM_SETTINGCHANGE;

    function ConvertImageToJpeg(const InputFileName, OutputFileName: string): Boolean;
  public
    { Public declarations }
    Directory: string;
    CurrentDir: string;
    CurrentFile: string;
    GitUrl: string;

    procedure Toast(aText, aTitle: string; sType: string = 'S,I,E'; ParentBase: TWinControl = nil);

    procedure populateCommands;
    procedure populateEnvironmentStrings;
    procedure populateMyFolders;
    procedure populateEnvExecutables;

    procedure UpdateTheme;
  end;

var
  Form1: TForm1;
  args: TStringList;

  function StartHook:BOOL; stdcall; external 'HotkeyHook.dll' name 'STARTHOOK';
  procedure StopHook; stdcall; external 'HotkeyHook.dll' name 'STOPHOOK';
  procedure SwitchToThisWindow(h1: hWnd; x: bool); stdcall;
  external user32 Name 'SwitchToThisWindow';

implementation

{$R *.dfm}

uses
  frmHover, UIAutomationClient, DarkModeApi.Vcl, Vcl.Themes,
  DarkModeApi, Winapi.UxTheme, CB.DarkMode, Ntapi.UserEnv, Ntapi.WinNt, Ntapi.ntrtl,
  pngimage, GIFImg, Cod.Imaging.Heif, Cod.Imaging.WebP, Vcl.SysStyles, ACL.Utils.Common;

type
  THostPreviewHandlerClass = class(THostPreviewHandler);

{ Global Functions}
function RtlGetVersion(var RTL_OSVERSIONINFOEXW): LONGINT; stdcall;
  external 'ntdll.dll' Name 'RtlGetVersion';
function isWindows11:Boolean;
var
  winver: RTL_OSVERSIONINFOEXW;
begin
  Result := False;
  if ((RtlGetVersion(winver) = 0) and (winver.dwMajorVersion>=10) and (winver.dwBuildNumber > 22000))  then
    Result := True;
end;

procedure EnableNCShadow(Wnd: HWND);
const
  DWMWCP_DEFAULT    = 0; // Let the system decide whether or not to round window corners
  DWMWCP_DONOTROUND = 1; // Never round window corners
  DWMWCP_ROUND      = 2; // Round the corners if appropriate
  DWMWCP_ROUNDSMALL = 3; // Round the corners if appropriate, with a small radius
  DWMWA_WINDOW_CORNER_PREFERENCE = 33; // [set] WINDOW_CORNER_PREFERENCE, Controls the policy that rounds top-level window corners
var
  DWM_WINDOW_CORNER_PREFERENCE: Cardinal;
begin

  if isWindows11  then
  begin

    DWM_WINDOW_CORNER_PREFERENCE := DWMWCP_ROUNDSMALL;
     DwmSetWindowAttribute(Wnd, DWMWA_WINDOW_CORNER_PREFERENCE, @DWM_WINDOW_CORNER_PREFERENCE, sizeof(DWM_WINDOW_CORNER_PREFERENCE));
  end;
end;


procedure UseImmersiveDarkMode(Handle: HWND; Enable: Boolean);
const
  DWMWA_USE_IMMERSIVE_DARK_MODE_BEFORE_20H1 = 19;
  DWMWA_USE_IMMERSIVE_DARK_MODE = 20;
var
  DarkMode: DWORD;
  Attribute: DWORD;
begin
//https://stackoverflow.com/a/62811758
  DarkMode := DWORD(Enable);

  if Win32MajorVersion = 10  then
  begin
    if Win32BuildNumber >= 17763 then
    begin
      Attribute := DWMWA_USE_IMMERSIVE_DARK_MODE_BEFORE_20H1;
    if Win32BuildNumber >= 18985 then
      Attribute := DWMWA_USE_IMMERSIVE_DARK_MODE;
      DwmSetWindowAttribute(Handle, Attribute, @DarkMode, SizeOf(DWord));
      SetWindowPos(Handle, 0, 0, 0, 0, 0, SWP_DRAWFRAME or SWP_NOACTIVATE or SWP_NOMOVE or SWP_NOSIZE or SWP_NOZORDER);
    end;
  end;
end;

function RunProcess(const Binary: string; const DirPath: string; args: TStrings): Boolean;
const
  BufSize = 4096; //1024
var
  Process: TProcess;
  Buf: AnsiString;
  Count: Integer;
  i: Integer;
  LineStart: Integer;
  OutputLine: AnsiString;
begin
  Process := TProcess.Create(nil);
  try
    Process.Executable := Binary;

    Process.Options := [poUsePipes, poStderrToOutPut];
    Process.ShowWindow := swoHIDE;

    Process.Parameters.Assign(args);
    Process.CurrentDirectory := DirPath;
    Process.Execute;

    OutputLine := '';
    SetLength(Buf, BufSize);
    repeat
      if (Process.Output <> nil) then
      begin
        Count := Process.Output.Read(PChar(Buf)^, BufSize);
      end
      else
        Count := 0;

      LineStart := 1;
      i := 1;
      while i <= Count do
      begin
        if CharInSet(Buf[i], [#10, #13]) then
        begin
          OutputLine := OutputLine + Copy(Buf, LineStart, i - LineStart);
          Form1.BCEditor1.Lines.Add(OutputLine);
          OutputLine := '';
          if (i < Count) and (CharInSet(Buf[i], [#10, #13])) and (Buf[i] <> Buf[i + 1]) then
            Inc(i);
          LineStart := i + 1;
        end;
        Inc(i);
      end;
      OutputLine := Copy(Buf, LineStart, Count - LineStart + 1);
    until Count = 0;

    if OutputLine <> '' then
      Form1.BCEditor1.Lines.Add(OutputLine);

    Process.WaitOnExit;
    Result := Process.ExitStatus = 0;
    if not Result then
      Form1.BCEditor1.Lines.Add('Command ' + Process.Executable + ' failed with exit code: ' + IntToStr(Process.ExitStatus));

  finally
    FreeAndNil(Process);
  end;
end;

function IsGitRepository(const Dir: string): Boolean;
var
  repo: Pgit_repository;
  dirPath: PAnsiChar;
  error: Integer;
begin
  dirPath := PAnsiChar(AnsiString(Dir));
  error := git_repository_open(@repo, dirPath);

  if error = 0 then
  begin
    git_repository_free(repo);
    Result := True;
  end
  else
    Result := False;
end;

function IsGit(const RepoDir): boolean;
var
  repo: Pgit_repository;
  remote: Pgit_remote;
  dirPath, remoteNamePAnsi: PAnsiChar;
  remoteURL: PAnsiChar;
  error: Integer;
begin
  Result := False;

  dirPath := PAnsiChar(AnsiString(RepoDir));

  // Open the repository
  error := git_repository_open(@repo, dirPath);
  if error <> 0 then
    Exit;
  Result := True;
  // Free the repository resource
  git_repository_free(repo);
end;


function GetRemoteURL(const RepoDir, RemoteName: string): string;
var
  repo: Pgit_repository;
  remote: Pgit_remote;
  dirPath, remoteNamePAnsi: PAnsiChar;
  remoteURL: PAnsiChar;
  error: Integer;
begin
  Result := '';

  dirPath := PAnsiChar(AnsiString(RepoDir));
  remoteNamePAnsi := PAnsiChar(AnsiString(RemoteName));

  // Open the repository
  error := git_repository_open(@repo, dirPath);
  if error <> 0 then
    Exit;

  // Look up the remote by its name
  error := git_remote_lookup(@remote, repo, remoteNamePAnsi);

  if error = 0 then
  begin
    // Get the remote URL
    remoteURL := git_remote_url(remote);
    Result := string(remoteURL);

    // Free the remote resource
    git_remote_free(remote);
  end;

  // Free the repository resource
  git_repository_free(repo);
end;

function ExtractThumbnail(Path: string; SizeX, SizeY: Integer; InitOle: Boolean = False): HBitmap;
var
  ShellFolder, DesktopShellFolder: IShellFolder;
  XtractImage: IExtractImage;
  Eaten: DWord;
  PIDL: PItemIDList;
  RunnableTask: IRunnableTask;
  Flags: DWord;
  Buf: array [0 .. MAX_PATH] of Char;
  BmpHandle: HBITMAP;
  Atribute, Priority: DWord;
  GetLocationRes: HResult;
  ASize: TSize;
begin
  Result := 0;
  try
    if InitOle then
      CoInitializeEx(nil, COINIT_APARTMENTTHREADED or COINIT_DISABLE_OLE1DDE);
    try
      OleCheck(SHGetDesktopFolder(DesktopShellFolder));
      OleCheck(DesktopShellFolder.ParseDisplayName(0, nil, StringToOleStr(ExtractFilePath(Path)),
          Eaten, PIDL, Atribute));
      OleCheck(DesktopShellFolder.BindToObject(PIDL, nil, IID_IShellFolder, Pointer(ShellFolder)));
      CoTaskMemFree(PIDL);

      OleCheck(ShellFolder.ParseDisplayName(0, nil, StringToOleStr(ExtractFileName(Path)), Eaten, PIDL, Atribute));
      ShellFolder.GetUIObjectOf(0, 1, PIDL, IExtractImage, nil, XtractImage);
      CoTaskMemFree(PIDL);

      if Assigned(XtractImage) then  // Try getting a thumbnail..
      begin
        RunnableTask := nil;
        ASize.cx := SizeX;
        ASize.cy := SizeY;
        Priority := 0;
        Flags:= IEIFLAG_ASPECT or IEIFLAG_OFFLINE or IEIFLAG_CACHE or IEIFLAG_QUALITY;
        GetLocationRes := XtractImage.GetLocation(Buf, SizeOf(Buf), Priority, ASize, 32, Flags);
        if (GetLocationRes = NOERROR) or (GetLocationRes = E_PENDING) then
        begin
          if GetLocationRes = E_PENDING then
            if XtractImage.QueryInterface(IRunnableTask, RunnableTask) <> S_OK then
              RunnableTask := nil;
          try
            //do not call OleCheck for debug
            XtractImage.Extract(BmpHandle);
            // This could consume a long time.
            Result := BmpHandle;
          except
            on E: EOleSysError do
              OutputDebugString(PChar(string(E.ClassName) + ': ' + E.message))
          end; // try/except
        end;
      end;

    finally
      if InitOle then
        CoUninitialize;
    end;
  except
    Result := 0;
  end;
end;

procedure HackAlpha(ABitmap: TBitmap; Color: TColor);
type
  PRGB32 = ^TRGB32;
  TRGB32 = record
    B, G, R, A: Byte;
  end;
  PPixel32 = ^TPixel32;
  TPixel32 = array[0..0] of TRGB32;
var
  Row: PPixel32;
  X, Y, slMain, slSize: Integer;
  R, G, B: Byte;
  c: Integer;
begin
  ABitmap.PixelFormat := pf32bit;
  c := ColorToRGB(Color);
  R := Byte(c);
  G := Byte(c shr 8);
  B := Byte(c shr 16);
  slMain := Integer(ABitmap.ScanLine[0]);
  slSize := Integer(ABitmap.ScanLine[1]) - slMain;
  for Y := 0 to ABitmap.Height - 1 do
  begin
    Row := PPixel32(slMain);
    for X := 0 to ABitmap.Width - 1 do
    begin
      Row[X].R := Row[X].A * (Row[X].R - R) shr 8 + R;
      Row[X].G := Row[X].A * (Row[X].G - G) shr 8 + G;
      Row[X].B := Row[X].A * (Row[X].B - B) shr 8 + B;
    end;
    slMain := slMain + slSize;
  end;
end;

function HackIconSize(ABitmap: TBitmap): TPoint;
type
  PPixel32 = ^TPixel32;
  TPixel32 = array [0..0] of Cardinal;
var
  Row: PPixel32;
  X, Y, i, j, slMain, slSize: Integer;
begin
  ABitmap.PixelFormat := pf32bit;
  Result.X := ABitmap.Width;
  Result.Y := ABitmap.Height;
  if (Result.X < 1) or (Result.Y < 1) then
    Exit;
  slMain := Integer(ABitmap.ScanLine[0]);
  slSize := Integer(ABitmap.ScanLine[1]) - slMain;
  Result.X := 0;
  Result.Y := 0;
  for Y := 0 to ABitmap.Height - 1 do
  begin
    Row := PPixel32(slMain);
    for X := 0 to ABitmap.Width - 1 do
    begin
      if (Row[X] and $FF000000) <> 0 then
      begin
        if X > Result.X then
          Result.X := X;
        if Y > Result.Y then
          Result.Y := Y;
      end;
    end;
    slMain := slMain + slSize;
  end;
  I := Math.Max(Result.X, Result.Y);
  j := 0;
  while I > j do
    j := j + 8;
  if j > 256 then
    j := 256;
  Result.X := j;
  Result.Y := Result.X;
end;

procedure GetIconFromFile(AFile: string; var AIcon: TIcon; SHIL_FLAG: Cardinal);
var
  LImgList: HIMAGELIST;
  SFI: TSHFileInfo;
  LIndex: Integer;
begin
  // Get the index of the imagelist
  SHGetFileInfo(PChar(AFile), FILE_ATTRIBUTE_NORMAL, SFI, SizeOf(TSHFileInfo),
    SHGFI_ICON {or SHGFI_LARGEICON} or SHGFI_SHELLICONSIZE or
    SHGFI_SYSICONINDEX or SHGFI_TYPENAME or SHGFI_DISPLAYNAME);
  if not Assigned(AIcon) then
    AIcon := TIcon.Create;
  // get image list
  SHGetImageList(SHIL_FLAG, IID_IImageList, Pointer(LImgList));
  // its index
  LIndex := SFI.iIcon;
  // seems that ILD_NORMAL returns bad result in Windows 7, so opt for ILD_IMAGE
  AIcon.Handle := ImageList_GetIcon(LImgList, LIndex, ILD_IMAGE);
end;

procedure Graphic2Bitmap(const ASrc: TGraphic; const ADest: TBitmap;
  const ATransparentColor: TColor);
var
  LCrop: TPoint;
begin
  if not Assigned(ASrc) or not Assigned(ADest) then
    Exit;
  if (ASrc.Width = 0) or (ASrc.Height = 0) then
    Exit;

  ADest.Width := ASrc.Width;
  ADest.Height := ASrc.Height;
  if ASrc.Transparent then
  begin
    ADest.Transparent := True;
    if (ATransparentColor <> clNone) then
    begin
      ADest.TransparentColor := ATransparentColor;
      ADest.TransparentMode := tmFixed;
      ADest.Canvas.Brush.Color := ATransparentColor;
    end
    else
      ADest.TransparentMode := tmAuto;
  end;

  ADest.Canvas.FillRect(Rect(0, 0, ADest.Width, ADest.Height));
  ADest.Canvas.Draw(0, 0, ASrc);
  LCrop := HackIconSize(ADest);
  ADest.Width := LCrop.X;
  ADest.Height := LCrop.Y;
end;

function Byte2Str(const i64Size: Int64): string;
const
  i64GB = 1024 * 1024 * 1024;
  i64MB = 1024 * 1024;
  i64KB = 1024;
begin
  if i64Size div i64GB > 0 then
    Result := Format('%.1f GB', [i64Size / i64GB])
  else if i64Size div i64MB > 0 then
    Result := Format('%.2f MB', [i64Size / i64MB])
  else if i64Size div i64KB > 0 then
    Result := Format('%.0f KB', [i64Size / i64KB])
  else
    Result := IntToStr(i64Size) + ' bytes';
end;

function CalcTHumbSize(Width, Height, ThumbWidth, ThumbHeight: Cardinal): Cardinal;
begin
  Result := 0;
  if (Width = 0) or (Height = 0) then
    Exit;
  if (Width < ThumbWidth) and (Height < ThumbHeight) then
    Result := (Width shl 16) + Height
  else
  begin
    if Width > Height then
    begin
      if Width < ThumbWidth then
        ThumbWidth := Width;
      Result := (ThumbWidth shl 16) + Trunc(ThumbWidth * Height / Width);
      if (Result and $FFFF) >ThumbHeight then
        Result := (Trunc(ThumbHeight * Width / Height) shl 16) + ThumbHeight;
    end
    else
    begin
      if Height < ThumbHeight then
        ThumbHeight := Height;
      Result := (Trunc(ThumbHeight * Width / Height) shl 16) + ThumbHeight;
      if ((Result shr 16) and $FFFF) > ThumbWidth then
        Result := (ThumbWidth shl 16) + Trunc(ThumbWidth * Height / Width);
    end;
  end;
end;

function Blend(Color1, Color2: TColor; A: Byte): TColor;
var
  C1, C2: LongInt;
  R, G, B, v1, v2: Byte;
begin
  A := Round(2.55 * A);
  C1 := ColorToRGB(Color1);
  C2 := COlorToRGB(COlor2);
  v1 := Byte(C1);
  v2 := Byte(C2);
  R := A * (v1 - v2) shr 8 + v2;
  v1 := Byte(C1 shr 8);
  v2 := Byte(C2 shr 8);
  G := A * (v1 - v2) shr 8 + v2;
  v1 := Byte(C1 shr 16);
  v2 := Byte(C2 shr 16);
  B := A * (v1 - v2) shr 8 + v2;
  Result := (B shl 16) + (G shl 8) + R;
end;

procedure WinGradient(DC: HDC; ARect: TRect; AColor1, AColor2: TColor);
var
  Vertexs: array[0..1] of TTriVertex;
  GRect: TGradientRect;
begin
  Vertexs[0].x := ARect.Left;
  Vertexs[0].y := ARect.Top;
  Vertexs[0].Red := (AColor1 and $000000FF) shl 8;
  Vertexs[0].Green := (AColor1 and $0000FF00);
  Vertexs[0].Blue := (AColor1 and $00FF0000) shr 8;
  Vertexs[0].Alpha := 0;
  Vertexs[1].x := ARect.Right;
  Vertexs[1].y := ARect.Bottom;
  Vertexs[1].Red := (AColor2 and $000000FF) shl 8;
  Vertexs[1].Green := (AColor2 and $0000FF00);
  Vertexs[1].Blue := (AColor2 and $00FF0000) shr 8;
  Vertexs[1].Alpha := 0;
  GRect.UpperLeft := 0;
  GRect.LowerRight := 1;
  GradientFill(DC, @Vertexs, 2, @GRect, 1, GRADIENT_FILL_RECT_V);
end;

function CompareNatural(s1, s2: string): Integer;
  function ExtractNr(n: Integer; var Txt: string): Int64;
  begin
    while (n <= Length(Txt)) and (Txt[n] >= '0') and (Txt[n] <= '9') do
      n := n + 1;
    Result := StrToInt64Def(Copy(Txt, 1, n - 1), 0);
    Delete(Txt, 1, (n - 1));
  end;
var
  B: Boolean;
begin
  Result := 0;
  s1 := LowerCase(s1);
  s2 := LowerCase(s2);
  if (s1 <> s2) and (s1 <> '') and (s2 <> '') then
  begin
    B := False;
    while (not B) do
    begin
      if ((s1[1] >= '0') and (s1[1] <= '9'))
      and ((s2[1] >= '0') and (s2[1] <= '9'))
      then
        Result := Sign(ExtractNr(1, s1) - ExtractNr(1, s2))
      else
        Result := Sign(Integer(s1[1]) - Integer(s2[1]));
      B := (Result <> 0) or (Min(Length(s1), Length(s2)) < 2);
      if not B then
      begin
        Delete(s1, 1, 1);
        Delete(s2, 1, 1);
      end;
    end;
  end;
  if Result = 0 then
  begin
    if (Length(s1) = 1) and (Length(s2) = 1) then
      Result := Sign(Integer(s1[1]) - Integer(s2[1]))
    else
      Result := Sign(Length(s1) - Length(s2));
  end;
end;

// a custom sort
function SortItem(List: rkIntegerList.TIntList; Index1, Index2: Integer): Integer;
var
  Item1, Item2: PItemData;
begin
  Item1 := Form1.Items[List[Index1]];
  Item2 := Form1.Items[List[Index2]];
  if Item1.Dir and Item2.Dir then
    Result := CompareNatural(Item1.Name, Item2.Name)
  else if Item1.Dir then
    Result := -1
  else if Item2.Dir then
    Result := 1
  else
    Result := CompareNatural(Item1.Name, Item2.Name);
end;

{ Form1 }

procedure TForm1.actPath2ClipExecute(Sender: TObject);
begin
// Copy current path to clipboard
  if not CurrentDir.IsEmpty and DirectoryExists(CurrentDir) then
  begin
    Clipboard.AsText := CurrentDir;
    Toast('Path copied to clipboard!', 'Current Path', 'S');
  end;
end;

procedure TForm1.actPreviewExecute(Sender: TObject);
begin
    pnlPreview.Visible := not pnlPreview.Visible;
end;

procedure TForm1.actSigIntExecute(Sender: TObject);
begin
  if DosCommand1.IsRunning then
    DosCommand1.SigInt;
end;

procedure TForm1.actUnPinExecute(Sender: TObject);
begin
  SpeedButton1Click(Sender);
end;

procedure TForm1.BCEditor1DblClick(Sender: TObject);
begin
  UpdateMainMenu(lastExplorerHandle);
end;

procedure TForm1.btnFileHandlerClick(Sender: TObject);
begin
  btnFileHandler.Visible := False;
  Panel1.Caption := '';
  if Assigned(fHexBuffer) then
    fHexBuffer.Free;
  ACLHexView1.Visible := False;
  try
  ACLHexView1.Data := nil;
  ACLHexView1.FullRefresh;
  except
  end;
end;

procedure TForm1.ButtonedEdit1Enter(Sender: TObject);
begin
// ButtonedEdit1.RightButton.Visible := True;
end;

procedure TForm1.ButtonedEdit1KeyPress(Sender: TObject; var Key: Char);
begin
// avoid ding sound
  if (Key = #13) or (Key = #27) then
    Key := #0;
end;

procedure TForm1.ButtonedEdit1KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var
  I: Integer;
  CLI: string;
begin
  CLI := ButtonedEdit1.Text;
  if key = 13 then
  begin
//    populateCommands;
    if CLI = 'list' then
    begin
      ListExplorerInstances;
      BCEditor1.Text := 'Current HWND: ' + IntToStr(lastExplorerHandle) + '';
      for I := 0 to lstExplorerPath.Count - 1 do
      begin
        if lstExplorerWnd[I] = IntToStr(lastExplorerHandle)  then

        BCEditor1.Text := BCEditor1.Text + #13#10 + lstExplorerPath[I] + ' ' + lstExplorerWnd[i];
      end;
    end
    else if CLI = 'items' then
    begin
      ListExplorerInstances;
      BCEditor1.Text := 'Current HWND: ' + IntToStr(lastExplorerHandle) + '';
      for I := 0 to lstExplorerItem.Count - 1 do
      begin
        BCEditor1.Text := BCEditor1.Text + #13#10 + lstExplorerItem[I];
      end;

    end
    else if CLI = '%' then
    begin
      populateEnvironmentStrings;
    end
    else if CLI = 'hexview' then
    begin
      var curFile := StatusBar1.Panels[0].Text;
      if FileExists(curFile) then
      begin
        if Assigned(fHexBuffer) then
          fHexBuffer.Free;
        fHexBuffer := TFileStream.Create(curFile, fmOpenRead);
        try
          Panel1.Caption := 'Hex: ' + curFile;
          btnFileHandler.Visible := True;
          ACLHexView1.Visible := True;
          ACLHexView1.StyleScrollBox.Reset;
          ACLHexView1.SetSelection(0, 0);
          ACLHexView1.Data := nil;
          ACLHexView1.FullRefresh;
          ACLHexView1.Data := fHexBuffer;
        finally
          //fHexBuffer.Free; //we should keep this open so the hex viewer will read on demand
        end;
      end;
    end
    else if CLI = 'preview' then
    begin
      var curFile := StatusBar1.Panels[0].Text;
      if FileExists(curFile) then
        BCEditor1.Lines.LoadFromFile(curFile);
    end
    else if CLI = 'tojpg' then
    begin
      var curFile := StatusBar1.Panels[0].Text;
      if FileExists(curFile) then
      begin
        if ConvertImageToJpeg(curFile, curFile +'.jpg') then
        begin
          BCEditor1.Clear;
          BCEditor1.Lines.Add('Image converted to JPG %90');
          BCEditor1.Lines.Add(curFile + '.jpg');
        end;
      end;
    end
    else if CLI = 'center' then
    begin
      if IsZoomed(lastExplorerHandle) then Exit;

      var _R: TRect;
      var _M: TMonitor;
      GetWindowRect(lastExplorerHandle, _R);
      _M := Screen.MonitorFromRect(_R);
      if (_R.Width > 0) and (_R.Height > 0)  then
      begin
        var NewPos: TPoint;
        NewPos.X := _M.Left + (_M.Width - _R.Width) div 2;
        NewPos.Y := _M.Top + (_M.Height - _R.Height) div 2;
        MoveWindow(lastExplorerHandle, NewPos.X, NewPos.Y, _R.Width, _R.Height, True);
      end;
    end
    else if CLI = 'cmd' then
    begin
      if DirectoryExists(lastExplorerPath) then
//        ShellExecute(0, PChar('OPEN'), PChar('cmd.exe'), PChar('/k refreshenv && cd /d ' + lastExplorerPath), PChar(lastExplorerPath), SW_SHOWNORMAL);
        ShellExecute(0, PChar('OPEN'), PChar('cmd.exe'), PChar('/k cd /d ' + lastExplorerPath), PChar(lastExplorerPath), SW_SHOWNORMAL)
      else
        ShellExecute(0, PChar('OPEN'), PChar('cmd.exe'), PChar('/k cd %USERPROFILE%'), nil, SW_SHOWNORMAL)
    end
    else if CLI = 'env' then
    begin
      BCEditor1.Lines.Clear;
      BCEditor1.Lines.Add('[Environment PATH]');
      for var _env in FEnvStrings do
        BCEditor1.Lines.Add(PChar(_env));
    end
    else if CLI = 'flushicons' then
    begin
      FlushIcons;
    end
    // show file explorer quick access directories
    else if CLI = ':' then
    begin
      populateMyFolders;
    end
    else if CLI = '>' then
    begin
      populateEnvExecutables;
    end
    else if Pos('>', CLI) = 1 then
    begin
      if Cli.Length > 1 then
      begin
        var command := Copy(CLI,2, Length(CLI) - 1);

        ShellExecute(0, PChar('OPEN'), PChar(command), nil, PChar(lastExplorerPath), SW_SHOWNORMAL);
      end
    end
    else if Pos('find ', CLI) = 1 then
    begin
      if DirectoryExists(lastExplorerPath) then
      begin
        DzDirSeek1.Dir := lastExplorerPath;
        DzDirSeek1.MaskKind := TDSMaskKind.mkInclusions;
        DzDirSeek1.Masks.Clear;
        DzDirSeek1.Masks.Add(Copy(CLI,6));
        DzDirSeek1.ResultKind := TDSResultKind.rkRelative;
        DzDirSeek1.Seek;
        BCEditor1.Lines.Clear;
        BCEditor1.Text := DzDirSeek1.List.GetText;
      end;
    end

    else if CLI = 'listexplorers' then
    begin
      ListBox1.Items := lstExplorerPath;
      ListBox1.Show;
      if ListBox1.Visible then
        ListBox1.SetFocus;
    end

    else if CLI = 'exit' then
      close
    else
    begin

      try
        begin
          BCEditor1.Lines.Clear;
          if CLI.Contains('=') then
          begin
            var ls := TStringList.Create;
            try
              ls.Delimiter := '=';
              ls.DelimitedText := CLI;
              if ls.Count > 1 then
              begin
                if DirectoryExists(ls[1]) then
                begin
                  ShellExecute(0, PChar('OPEN'), PChar(ls[1]), nil, nil, SW_SHOWNORMAL);
                end;
              end;

            finally
              ls.Free;
            end;
          end
          else
          if DirectoryExists(lastExplorerPath) then
          begin
            var basePath := lastExplorerPath;
            if DirectoryExists(CurrentFile) then
              basePath := CurrentFile;

// Temporary disabled to try DOSCommand Instead
//            args := TStringList.Create;
//            args.Add('/c');
////            args.Add('chcp');
////            args.Add('65001');
////            args.Add('&');

            if OpenURL1.Enabled then //git folder detected
            begin
              if (CLI = 'gp') or CLI.Contains('-pull') then
              begin
                ButtonedEdit1.Text := 'git -c fetch.parallel=0 -c submodule.fetchjobs=0 pull --progress "origin"';
              end
              else if (CLI = 'gu') or CLI.Contains('-url') then
              begin
                ButtonedEdit1.Text := 'giturl';
              end
              else if (CLI = 'gr') or Cli.Contains('-readme') then
              begin
                var readmePath := basePath + '\README.md';
                if FileExists(readmePath) then
                  ButtonedEdit1.Text := ('start ' + readmePath)
                else
                  ButtonedEdit1.Text := ('echo NO README FOUND!');
              end;
              CLI := ButtonedEdit1.Text;
            end;
//            args.Add(CLI);
//            RunProcess('cmd.exe', PChar(basePath), args);
//            Toast('Command finished!', '','S');
//            BCEditor1.GotoLineAndCenter(BCEditor1.Lines.Count);
//            args.Free;
//            args := nil;

            DosCommand1.CurrentDir := basePath;
    //        DosCommand1.CommandLine := 'cmd.exe /c ' + ButtonedEdit1.Text;
    //        DosCommand1.Execute;
            ProcessDosCommand(Self, PChar('cmd.exe /c ' + CLI));
          end;
        end
      except
        //on E:Exception do

      end;

    end;
    ButtonedEdit1.Text := '';
 end;
end;

procedure TForm1.CMFocusChanged(var Msg: TCMFocusChanged);
begin
  ListBox1.Visible := ListBox1.Focused;


  inherited;
end;

function TForm1.ConvertImageToJpeg(const InputFileName,
  OutputFileName: string): Boolean;
const
  // Hex headers for different formats
  BMPHeader: array[0..1] of Byte = ($42, $4D); // BM
  PNGHeader: array[0..7] of Byte = ($89, $50, $4E, $47, $0D, $0A, $1A, $0A); // PNG signature
  GIFHeader: array[0..2] of Byte = ($47, $49, $46); // GIF
  JPEGHeader: array[0..1] of Byte = ($FF, $D8); // JPEG SOI marker
  WebPHeader: array[0..3] of Byte = ($52, $49, $46, $46); // RIFF for WebP
  HEIFHeader: array[0..3] of Byte = ($66, $74, $79, $70); // ftyp for HEIF
var
  FileStream: TFileStream;
  Header: TBytes;
  InputImage: TGraphic;
  JPEGImage: TJPEGImage;
  Bitmap: TBitmap;
  FormatValid: Boolean;

  function CompareHeader(const FileHeader, ValidHeader: array of Byte): Boolean;
  var
    I: Integer;
  begin
    Result := Length(FileHeader) >= Length(ValidHeader);
    if Result then
      for I := 0 to High(ValidHeader) do
        if FileHeader[I] <> ValidHeader[I] then
          Exit(False);
  end;

  function ConfirmOverwrite(const FileName: string): Boolean;
  begin
    Result := not FileExists(FileName) or
      (MessageDlg(Format('File "%s" already exists. Do you want to overwrite it?',
        [FileName]), mtConfirmation, [mbYes, mbNo], 0) = mrYes);
  end;

begin
  Result := False;
  FormatValid := False;
  Header := nil;
  InputImage := nil;
  JPEGImage := nil;
  Bitmap := nil;

  // Check if output file exists and confirm overwrite
  if not ConfirmOverwrite(OutputFileName) then
    Exit;

  try
    // Open file to read header
    FileStream := TFileStream.Create(InputFileName, fmOpenRead or fmShareDenyWrite);
    try
      SetLength(Header, 8); // Longest header is 8 bytes (PNG)
      FileStream.ReadBuffer(Header[0], Length(Header));
    finally
      FileStream.Free;
    end;

    // Validate format based on header
    if CompareHeader(Header, BMPHeader) then
      InputImage := TBitmap.Create
    else if CompareHeader(Header, PNGHeader) then
      InputImage := TPngImage.Create
    else if CompareHeader(Header, GIFHeader) then
      InputImage := TGIFImage.Create
    else if CompareHeader(Header, JPEGHeader) then
      InputImage := TJPEGImage.Create
    else if CompareHeader(Header, WebPHeader) then
      InputImage := TWebPImage.Create
    else if CompareHeader(Header, HEIFHeader) then
      InputImage := THEIFImage.Create
    else
      raise EInvalidImageFormat.Create('Unsupported image format.');

    FormatValid := True;

    // Load image into InputImage
    InputImage.LoadFromFile(InputFileName);

    // Create intermediate bitmap for PNG and HEIF
    if (InputImage is TPngImage) or (InputImage is THEIFImage) then
    begin
      Bitmap := TBitmap.Create;
      try
        Bitmap.Width := InputImage.Width;
        Bitmap.Height := InputImage.Height;
        Bitmap.Canvas.Draw(0, 0, InputImage);

        // Convert to JPEG
        JPEGImage := TJPEGImage.Create;
        try
          JPEGImage.Assign(Bitmap); // Assign from bitmap instead of direct conversion
          JPEGImage.CompressionQuality := 90;
          JPEGImage.SaveToFile(OutputFileName);
          Result := True;
        finally
          JPEGImage.Free;
        end;
      finally
        Bitmap.Free;
      end;
    end
    else
    begin
      // Direct conversion for other formats
      JPEGImage := TJPEGImage.Create;
      try
        JPEGImage.Assign(InputImage);
        JPEGImage.CompressionQuality := 90;
        JPEGImage.SaveToFile(OutputFileName);
        Result := True;
      finally
        JPEGImage.Free;
      end;
    end;
  except
    on E: Exception do
      raise Exception.CreateFmt('Error converting image: %s', [E.Message]);
  end;

  // Clean up
  if not FormatValid then
    raise EInvalidImageFormat.Create('Image format validation failed.');
  if Assigned(InputImage) then
    InputImage.Free;
end;

procedure TForm1.CopyPathtoClipboard1Click(Sender: TObject);
begin
  Clipboard.SetTextBuf(PChar(CurrentDir));
end;

procedure TForm1.CreateParams(var Params: TCreateParams);
begin
  inherited;

  Params.WinClassName := 'ExplorerCommandWnd';
end;

procedure TForm1.DosCommand1ExecuteError(ASender: TObject; AE: Exception;
  var AHandled: Boolean);
begin
  if AHandled then
    BCEditor1.Lines.Text := AE.ToString;
end;

//procedure TForm1.DosCommand1NewChar(ASender: TObject; ANewChar: Char);
//begin
//  BCEditor1.BeginUpdate;
//
//  if ANewChar <> #13 then
//  BCEditor1.Text := BCEditor1.Text + ANewChar;
//
//  BCEditor1.GotoLineAndCenter(BCEditor1.Lines.Count);
//
//  BCEditor1.EndUpdate;
//  KHexEditor1.ExecuteCommand(ecInsertString, PChar(ANewChar));
//end;

procedure TForm1.DosCommand1NewLine(ASender: TObject; const ANewLine: string;
  AOutputType: TOutputType);
begin
////  AOutputType := otEntireLine;
////  BCEditor1.Lines.Add(ANewLine);
////  BCEditor1.Text := BCEditor1.Text +#13#10+ ANewLine;
//  FCommandOutput.Add(ANewLine);
//  BCEditor1.BeginUpdate;
//  BCEditor1.Lines :=  FCommandOutput;
////  KHexEditor1.ExecuteCommand(ecInsertString, PChar(ANewLine));
////  BCEditor1.Perform(EM_SCROLL, SB_LINEDOWN, 0);
//  BCEditor1.GotoLineAndCenter(BCEditor1.Lines.Count);
//  BCEditor1.EndUpdate;
//  Application.ProcessMessages;

  FSyncLock.Enter;
  try
    FOutputBuffer.Add(ANewLine);
  finally
    FSyncLock.Leave;
  end;
end;

procedure TForm1.DosCommand1Terminated(Sender: TObject);
begin
  BCEditor1.Lines.Add('¡Completed process!');
  ActivityIndicator1.Animate := False;
  ActivityIndicator1.Visible := False;
//  BCEditor1.Lines := FCommandOutput;
  BCEditor1.GotoLineAndCenter(BCEditor1.Lines.Count);
  FCommandOutput.Clear;
end;

procedure TForm1.DosCommand1TerminateProcess(ASender: TObject;
  var ACanTerminate: Boolean);
begin
  ACanTerminate := True;
end;

procedure TForm1.Exit1Click(Sender: TObject);
begin
  Close;
end;

procedure TForm1.FlushIcons;
var
  DesktopFolder: IShellFolder;
  Pidl: PItemIDList;
begin
  // Get the desktop folder
  SHGetDesktopFolder(DesktopFolder);

  // Get the PIDL for the desktop folder
  SHGetSpecialFolderLocation(0, CSIDL_DESKTOP, Pidl);

  try
    // Notify the system of the association change
    SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, Pidl, nil);
  finally
    // Free the PIDL
    CoTaskMemFree(Pidl);
  end;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin

  AllowSetForegroundWindow(GetCurrentProcessId);

  if not StartHook then
  begin
    MessageDlg('Couldn''t set global hotkey.',mtError, [mbOK], 0);
    Application.Terminate;
  end;

  KeyPreview := True;

  ActivityIndicator1.Visible := False;

  lstExplorerPath := TStringList.Create;
  lstExplorerWnd := TStringList.Create;
  lstExplorerItem := TStringList.Create;

  Application.OnDeactivate := OnFocusLost;

  // IShellPreview
  fPreview := nil;

//  BCEditor1.Font.Name := 'Consolas';
//  BCEditor1.Font.Size := 9;

  FEnvExecutables := TStringList.Create;
  FEnvStrings := TStringList.Create;

  RefreshEnvironmentVariables;

  // IAutoComplete
  ButtonedEdit1.ACEnabled := True;
  ButtonedEdit1.ACOptions := [acAutoAppend, acAutoSuggest, acUpDownKeyDropsList];
  ButtonedEdit1.ACSource := acsList;
  populateCommands;

  FCommandOutput := TStringList.Create;

  // LibGit2 initialization
//  git_libgit2_init;
  InitLibgit2;

  mnuAutoStart.Checked := AppAutoStart1.IsStartupEnabled;

  // Speeding up the DOSCommand output
  FOutputBuffer := TStringList.Create;
  FSyncLock := TCriticalSection.Create;

//  SetWindowColorModeAsSystem;
  UpdateTheme;
end;

procedure TForm1.FormDestroy(Sender: TObject);
begin
  FSyncLock.Free;
  FOutputBuffer.Free;

  ShutdownLibgit2;
//  git_libgit2_shutdown;

  FEnvStrings.Free;
  FEnvExecutables.Free;


  DosCommand1.Stop;
  FCommandOutput.Free;

  if fPreview <> nil then
    fPreview.Free;

  lstExplorerWnd.Free;
  lstExplorerPath.Free;
  lstExplorerItem.Free;

  StopHook;
end;

procedure TForm1.FormKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if (key = VK_ESCAPE) and not FPinned then
    Hide;
end;

procedure TForm1.FormShow(Sender: TObject);
begin
  ShowWindow(Application.Handle, SW_HIDE);

//  if GetForegroundWindow <> Handle then
//  begin
//    SwitchToThisWindow(Handle, );
//    SetForegroundWindow(Handle);
//  end;

end;

// since Windows 11 22h2 build 22621.2506 new file explorer address bar is available
// so it should choose the new address bar
function TForm1.GetExplorerAddressBarRect(AHandle: HWND): TRect;
var
  ExplorerRect: TRect;
  LWND: HWND;
begin
  // we assume it is a valid explorer instance before calling this function
  Winapi.Windows.GetWindowRect(AHandle, ExplorerRect);

  LWND := FindWindowEx(AHandle, 0, 'WorkerW', nil);
  if LWND > 0 then
    LWND := FindWindowEx(LWND, 0, 'ReBarWindow32', nil);
  if LWND > 0 then
    LWND := FindWindowEx(LWND, 0, 'Address Band Root', nil);
  if LWND > 0 then
    LWND := FindWindowEx(LWND, 0, 'msctls_progress32', nil);
  if LWND > 0 then
    LWND := FindWindowEx(LWND, 0, 'Breadcrumb Parent', nil);
  if LWND > 0 then
  begin
    Winapi.Windows.GetWindowRect(LWND, Result);
//    Result.Width := Width;
    if Result.Width < 600 then
      Result.Width := 600;
    Result.Height := Height;
  end
  else
  begin
    // on newer File Explorer on Windows 11 let's pick the empty area of the
    // Child Class: Microsoft.UI.Content.DesktopChildSiteBridge (top area)
    LWND := FindWindowEx(AHandle, 0, 'Microsoft.UI.Content.DesktopChildSiteBridge', nil);
    if LWND > 0 then
    begin
      var nRect: TRect;
      Winapi.Windows.GetWindowRect(LWND, nRect);
      Result.Width := Width;
      Result.Height := Height;
      Result.Left := ExplorerRect.Left + (ExplorerRect.Width - Result.Width) div 2;
      Result.Top := ExplorerRect.Top + nRect.Height;
    end
    else
    begin
      // it might be a different explorer version, maybe the newer on Windows 11 Insider which changed its address bar position
      Result.Width := Width;
      Result.Height := Height;
      Result.Left := ExplorerRect.Left + (ExplorerRect.Width - Result.Width) div 2;
      Result.Top := ExplorerRect.Top + (ExplorerRect.Height - Height) div 2;
    end;
  end;
end;

procedure TForm1.KeyEventHandler(var Msg: TMessage);
var
  I: Integer;
  command: String;
  Ret: Integer;

  HActiveWindow: HWND;
  HForegroundThread, HAppThread: DWORD;
  FClientId: DWORD;
  Win11TabContainer: HWND; //TITLE_BAR_SCAFFOLDING_WINDOW_CLASS
begin
    populateCommands;
//  OutputDebugString(PChar('heehhehe'));
  command := PChar(Msg.LParam);
  lastExplorerHandle := StrToInt(command);
  lastExplorerPath := '';

  if not Visible then
  begin
    var rct: TRect;
    rct := GetExplorerAddressBarRect(lastExplorerHandle);
    Left := rct.Left;
    Width := rct.Width;
    if Width < 800 then
      Width := 800;
    Top := rct.Top;

//    SwitchToThisWindow(GetDesktopWindow, True);

//    BorderStyle := bsNone;
//    AnimateWindow(Handle, 128, AW_SLIDE or AW_VER_POSITIVE );
//    BorderStyle := bsSizeable;
    Show;

    HActiveWindow := GetForegroundWindow();
    HForegroundThread := GetWindowThreadProcessId(HActiveWindow, @FClientId);
    AllowSetForegroundWindow(FClientId);

    HAppThread := GetCurrentThreadId;

    if not SetForegroundWindow(Handle) then
      SwitchToThisWindow(GetDesktopWindow, True);

    // magic part to switch correctly to our window
    if HForegroundThread <> HAppThread then
    begin
      AttachThreadInput(HForegroundThread, HAppThread, True);
      BringWindowToTop(Handle);
      Winapi.Windows.SetFocus(Handle);
      AttachThreadInput(HForegroundThread, HAppThread, False);
    end;

//    Winapi.Windows.GetWindowRect(HActiveWindow, rct);
    SetWindowPos(Handle, HWND_TOP, 0, 0, 0, 0, {SWP_ASYNCWINDOWPOS or }SWP_NOMOVE or SWP_NOSIZE or SWP_SHOWWINDOW);

    // let's put out menu in the Explorer window
//    Win11TabContainer := FindWindowEx(HActiveWindow, 0, 'TITLE_BAR_SCAFFOLDING_WINDOW_CLASS', nil);
//    if Win11TabContainer >  0 then
//    begin
//      formHover.Show;
//      var mr: TRect;
//      Winapi.Windows.GetWindowRect(Win11TabContainer, mr);
//      formHover.Left := 0;
//      formHover.Top := 0;
//      formHover.Width := mr.Width;
//      formHover.Height := mr.Height;
//      formHover.BoundsRect := mr;
//      Winapi.Windows.SetParent(formHover.Handle, Win11TabContainer);
//    end;


//    ButtonedEdit1.SetFocus;

    // before listing explorer instances let's see if we are in a open save dialog
    // WorkerW->ReBarWindow32->Address Band Root->msctls_progress32->ComboBoxEx32->ComboBox->Edit

    Ret := ListExplorerInstances;

{    var FirstPath := IntToStr(Ret);
    StatusBar1.Panels[0].Text := FirstPath;
    }

    for I := 0 to lstExplorerWnd.Count - 1 do
    begin
      if lstExplorerWnd[i] = IntToStr(lastExplorerHandle) then
      begin
        lastExplorerPath := lstExplorerPath[I];
        //rkSmartPath1.Path := lstExplorerPath[I];
        StatusBar1.Panels[0].Text := lstExplorerItem[i];
        //WIC
        if FileExists(lstExplorerItem[i]) then
          ShowPreview(lstExplorerItem[i]);
        CurrentFile := lstExplorerItem[i];
      end;
    end;

    if DirectoryExists(StatusBar1.Panels[0].Text) then
      rkSmartPath1.Path := StatusBar1.Panels[0].Text
    else if DirectoryExists(ExtractFilePath(StatusBar1.Panels[0].Text)) then
    begin
         rkSmartPath1.Path := ExtractFilePath(StatusBar1.Panels[0].Text);
    end;

    if IsGitRepository(rkSmartPath1.Path) then
    begin
      ButtonedEdit1.LeftButton.ImageIndex := 3
    end
    else
      ButtonedEdit1.LeftButton.ImageIndex := 0;

    CurrentDir := rkSmartPath1.Path;
    GitUrl := GetRemoteURL(rkSmartPath1.Path, 'origin');
    if Pos('http', LowerCase(GitUrl)) = 1 then
    begin
      OpenURL1.Enabled := True;
      pnlTitle.Visible := True;
      LinkLabel1.Caption := 'Repository: <a href="' +  GitUrl + '">' + GitUrl + '</a>';
      LinkLabel1.Left := (pnlTitle.Width - LinkLabel1.Width) div 2;
    end
    else
    begin
      OpenURL1.Enabled := False;
      pnlTitle.Visible := False;
    end;

//    BCEditor1.Lines.Add(gurl);
  end
  else
  begin
//    SwitchToThisWindow(Handle, True);
    Hide;
  end;
end;

procedure TForm1.KeyEventHandlerAll(var Msg: TMessage);
var
  I: Integer;
  command: String;
  Ret: Integer;

  HActiveWindow: HWND;
  HForegroundThread, HAppThread: DWORD;
  FClientId: DWORD;

begin
//  OutputDebugString(PChar('heehhehe'));
  command := PChar(Msg.LParam);
  lastExplorerHandle := StrToInt(command);
  lastExplorerPath := '';

  if not Visible then
  begin
//    SwitchToThisWindow(GetDesktopWindow, True);
    Show;
    HActiveWindow := GetForegroundWindow();
//    UpdateMainMenu(lastExplorerHandle);
    HForegroundThread := GetWindowThreadProcessId(HActiveWindow, @FClientId);
    AllowSetForegroundWindow(FClientId);

    HAppThread := GetCurrentThreadId;

    if not SetForegroundWindow(Handle) then
      SwitchToThisWindow(GetDesktopWindow, True);



    // magic part to switch correctly to our window
    if HForegroundThread <> HAppThread then
    begin
      AttachThreadInput(HForegroundThread, HAppThread, True);
      BringWindowToTop(Handle);
      Winapi.Windows.SetFocus(Handle);
      AttachThreadInput(HForegroundThread, HAppThread, False);
    end;

    var rct: TRect;
    Winapi.Windows.GetWindowRect(HActiveWindow, rct);
    SetWindowPos(Handle, HWND_TOP, 0, 0, 0, 0, {SWP_ASYNCWINDOWPOS or }SWP_NOMOVE or SWP_NOSIZE or SWP_SHOWWINDOW);
    // center to current window, otherwise to monitors
    var nLeft := rct.Left + (rct.Width - Width) div 2;
    var nTop := rct.Top + (rct.Height - Height) div 2;
    if nLeft < 0 then
      Left := (Screen.Width - Width) div 2
    else
      Left := nLeft;
    if nTop < 0 then
      Top := (Screen.Height - Height) div 2
    else
      Top := nTop;

//    ButtonedEdit1.SetFocus;

    // before listing explorer instances let's see if we are in a open save dialog
    // WorkerW->ReBarWindow32->Address Band Root->msctls_progress32->ComboBoxEx32->ComboBox->Edit

    Ret := ListExplorerInstances;

{    var FirstPath := IntToStr(Ret);
    StatusBar1.Panels[0].Text := FirstPath;
    }

    for I := 0 to lstExplorerWnd.Count - 1 do
    begin
      if lstExplorerWnd[i] = IntToStr(lastExplorerHandle) then
      begin
        lastExplorerPath := lstExplorerPath[I];
        //rkSmartPath1.Path := lstExplorerPath[I];
        StatusBar1.Panels[0].Text := lstExplorerItem[i];
        //WIC
        if FileExists(lstExplorerItem[i]) then
          ShowPreview(lstExplorerItem[i]);
        CurrentFile := lstExplorerItem[i];
      end;
    end;

    if DirectoryExists(StatusBar1.Panels[0].Text) then
      rkSmartPath1.Path := StatusBar1.Panels[0].Text
    else if DirectoryExists(ExtractFilePath(StatusBar1.Panels[0].Text)) then
    begin
         rkSmartPath1.Path := ExtractFilePath(StatusBar1.Panels[0].Text);
    end;

    if IsGitRepository(rkSmartPath1.Path) then
    begin
      ButtonedEdit1.LeftButton.ImageIndex := 3
    end
    else
      ButtonedEdit1.LeftButton.ImageIndex := 0;

    CurrentDir := rkSmartPath1.Path;
    GitUrl := GetRemoteURL(rkSmartPath1.Path, 'origin');
    if Pos('http', LowerCase(GitUrl)) = 1 then
      OpenURL1.Enabled := True
    else
      OpenURL1.Enabled := False;

//    BCEditor1.Lines.Add(gurl);
  end
  else
  begin
//    SwitchToThisWindow(Handle, True);
    Hide;
  end;
end;

// Lists explorer instances which has items visible, ignores special directories
procedure TForm1.LinkLabel1LinkClick(Sender: TObject; const Link: string;
  LinkType: TSysLinkType);
begin
  if LinkType = sltURL then
  begin
    ShellExecute(0, 'OPEN', PChar(Link), nil, nil, SW_NORMAL);
  end;
end;

procedure TForm1.ListBox1DblClick(Sender: TObject);
begin
  Hide;
  SwitchToThisWindow(StrToInt(lstExplorerWnd[ListBox1.ItemIndex]), True);
end;

procedure TForm1.ListBox1KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = 13 then
  begin
    Hide;
//    SwitchToThisWindow(StrToInt(lstExplorerWnd[ListBox1.ItemIndex]), True);
    SwitchToWindow(StrToInt(lstExplorerWnd[ListBox1.ItemIndex]));
    ListBox1.Visible := False;
  end;
end;

function GetSpecialFolderPath(const FolderID: Integer): string;
var
  ShellFolder: IShellFolder;
  IDList: PItemIDList;
  StrRet: TStrRet;
  FolderPath: array [0..MAX_PATH] of Char;
begin
  Result := '';

  if Succeeded(SHGetSpecialFolderLocation(0, FolderID, IDList)) then
  begin
    if Succeeded(SHGetDesktopFolder(ShellFolder)) then
    begin
      if Succeeded(ShellFolder.GetDisplayNameOf(IDList, SHGDN_FORPARSING, StrRet)) then
      begin
        if StrRet.uType = STRRET_WSTR then
        begin
          OleStrToStrVar(StrRet.pOleStr, Result);
          CoTaskMemFree(StrRet.pOleStr);
        end
        else
        begin
        // FIXLATER
//          StrRetToStr(StrRet, IDList, FolderPath, SizeOf(FolderPath));
          Result := FolderPath;
        end;
      end;
    end;

    CoTaskMemFree(IDList);
  end;
end;

function TForm1.ListExplorerInstances: Integer;
const
  IID_IServiceProvider: TGUID = '{6D5140C1-7436-11CE-8034-00AA006009FA}';
  SID_STopLevelBrowser: TGUID = '{4C96BE40-915C-11CF-99D3-00AA004AE837}';
  LOOPTIME = 500; //ms i.e. half a second
var
  ShellWindows: IShellWindows;
  I: Integer;
  ShellBrowser: IShellBrowser;
  WndIface: IDispatch;
  WebBrowserApp: IWebBrowserApp;
  ServiceProvider: IServiceProvider;
  ItemIDList, ItemIDList2: PItemIDList;
  bar: HWND;
  ShellView: IShellView;
  FolderView: IFolderView;
  PersistFolder2: IPersistFolder2;
  ShellFolder: IShellFolder;
  focus: Integer;
  ret: _STRRET;
  folderPath: PChar;
  AMalloc: IMalloc;
  hr: HRESULT;
  CurTime: Int64;
  // Thumbnail
//  ItemIDList3: PItemIDList;
//  Thumbnail: IExtractImage;
//  ThumbBuf: array[0..MAX_PATH] of Char;
//  Runnable: IRunnableTask;
//  Flags, Priority: DWORD;
//  BmpHandle: HBITMAP;
//  ASize: TSize;
//  GetLocationRes: HRESULT;
begin
  Result := 0;
  lstExplorerPath.BeginUpdate;
  lstExplorerPath.Clear;
  lstExplorerWnd.BeginUpdate;
  lstExplorerWnd.Clear;
  lstExplorerItem.BeginUpdate;
  lstExplorerItem.Clear;

  hr := CoInitializeEx(nil, COINIT_APARTMENTTHREADED); // <-- manually call CoInitialize()
  if Succeeded(hr) then
  begin

  // this might fail on first try, so let's insist for LOOPTIME ms
    hr := CoCreateInstance(CLASS_ShellWindows, nil, CLSCTX_ALL,
      IID_IShellWindows, ShellWindows);
    CurTime := GetTickCount64;
    while not Succeeded(hr) do
    begin
      if ((GetTickCount64-CurTime)>LOOPTIME) then break;
      hr := CoCreateInstance(CLASS_ShellWindows, nil, CLSCTX_ALL,
        IID_IShellWindows, ShellWindows);
    end;

    if Succeeded(hr) then
    begin
      Result := 1;
      for I := 0 to ShellWindows.Count - 1 do
      begin
        if VarType(ShellWindows.Item(I)) = varDispatch then
        begin
          WndIface := ShellWindows.Item(VarAsType(I, VT_I4));
          if WndIface <> nil then
          try
          if Succeeded(WndIface.QueryInterface(IID_IWebBrowserApp, WebBrowserApp)) then
          begin
            lstExplorerWnd.Add(inttostr(WebBrowserApp.HWND));

            begin
              if Succeeded(WebBrowserApp.QueryInterface(IID_IServiceProvider,
                ServiceProvider)) then
              begin
                if Succeeded(ServiceProvider.QueryService(SID_STopLevelBrowser,
                  IID_IShellBrowser, ShellBrowser)) then
                begin
                  if Succeeded(ShellBrowser.QueryActiveShellView(ShellView)) then
                  begin
                    if Succeeded(ShellView.QueryInterface(IID_IFolderView, FolderView)) then
                    begin
                      FolderView.GetFocusedItem(focus);
                      FolderView.Item(focus,ItemIDList);
                      if Succeeded(FolderView.GetFolder(IID_IPersistFolder2, PersistFolder2)) then
                      begin
                        if succeeded(PersistFolder2.GetCurFolder(ItemIDList2)) then
                        begin
                          // mmmm
{                          if (ItemIDList <> nil)
                          and Succeeded(ShellFolder.GetDisplayNameOf(ItemIDList, SHGDN_FORPARSING, Ret))
                           then
                          begin
                            case Ret.uType of
                              STRRET_WSTR:
                              begin
//                                FolderPath := StrPas(Ret.pOleStr);
                                CoTaskMemFree(Ret.pOleStr);
                              end;
                              STRRET_CSTR:
                              begin
//                                FolderPath := Ret.cStr;
                              end;
                              STRRET_OFFSET:
                              begin
                                FolderPath := PChar(Integer(ItemIDList) + Ret.uOffset);
                              end
                              else
                                FolderPath := ' ';
                            end;
                          end;}

                          folderPath := StrAlloc(MAX_PATH);
                          if SHGetPathFromIDList(ItemIDList2, folderPath) then
                            lstExplorerPath.Add(folderPath);
                          SHGetMalloc(AMalloc);
                          AMalloc.Free(ItemIDList2);
                          StrDispose(folderPath);
                        end;

                        if Succeeded(PersistFolder2.QueryInterface(IID_IShellFolder, ShellFolder)) then
                        begin
                          if (ItemIDList <> nil) and Succeeded(ShellFolder.GetDisplayNameOf(ItemIDList, SHGDN_FORPARSING, ret)) then
                            lstExplorerItem.Add(ret.pOleStr)
                          else
                            lstExplorerItem.Add('no name');
                        end;

//                        //extract thumbnail
//                        if Succeeded(ShellFolder.GetUIObjectOf(0, 1, ItemIDList3, IExtractImage, nil, Thumbnail)) then
//                        begin
//                          CoTaskMemFree(ItemIDList3);
//
//                          if Assigned(Thumbnail) then
//                          begin
//                            Runnable := nil;
//                            ASize.cx := 256;
//                            ASize.cy := 256;
//                            Priority := 0;
//                            Flags := IEIFLAG_ASPECT or IEIFLAG_OFFLINE or IEIFLAG_CACHE or IEIFLAG_QUALITY;
//                            GetLocationRes := Thumbnail.GetLocation(ThumbBuf, SizeOf(ThumbBuf), Priority, ASize, 32, Flags);
//                            if (GetLocationRes = NOERROR) or (GetLocationRes = E_PENDING) then
//                            begin
//                              if GetLocationRes = E_PENDING then
//                                if Thumbnail.QueryInterface(IRunnableTask, Runnable) <> S_OK then
//                                  Runnable := nil;
//                              try
//                                Thumbnail.Extract(BmpHandle);
//                                Image1.Picture.Bitmap.Handle := BmpHandle;
//                              except
//                                on E: EOleSysError do
//                                  OutputDebugString(PChar(string(E.ClassName) + ': ' + E.Message));
//                              end;
//                            end;
//                          end;
//                        end;
                      end;
                    end;
                  end;

                end;

              end;
            end;
            // make sure the other lists are even to found explorer
            if lstExplorerWnd.Count > lstExplorerPath.Count then
              lstExplorerPath.Add('');
            if lstExplorerWnd.Count > lstExplorerItem.Count then
              lstExplorerPath.Add('');
          end;
          except
          end;
        end;

      end;

    end;
  end;
  CoUninitialize; // <-- free memory
  lstExplorerItem.EndUpdate;
  lstExplorerWnd.EndUpdate;
  lstExplorerPath.EndUpdate;
end;

procedure TForm1.mnuAutoStartClick(Sender: TObject);
begin
  mnuAutoStart.Checked := not mnuAutoStart.Checked;
  AppAutoStart1.Enabled := mnuAutoStart.Checked;
end;

procedure TForm1.NoBorder(var Msg: TWMNCActivate);
begin
  Msg.Active := False;
  inherited;
end;

procedure TForm1.OnFocusLost(Sender: TObject);
begin
  EsImage1.Picture.Assign(nil);

  StatusBar1.Panels[0].Text := '';
  if not FPinned then
  Hide;
end;

procedure TForm1.OpenURL1Click(Sender: TObject);
begin
  ShellExecute(0, 'OPEN', PChar(GitUrl), nil, nil, SW_SHOWNORMAL);
end;

procedure TForm1.populateCommands;
begin
  FCommandType := ctNormal;

  with ButtonedEdit1.ACStrings do
  begin
    BeginUpdate;
    Clear;
    Add('>');
    Add(':');
    Add('%');
    Add('help');
    Add('exit');
    Add('find');
    Add('open');
    Add('cmd');
    Add('env');
    Add('dir');
    if OpenURL1.Enabled then
    begin
      Add('git');
      Add('git-pull'); // git pull
      Add('gp');
      Add('git-readme'); // git readme
      Add('gr');
      Add('git-url'); // git url
      Add('gu');
    end;
    Add('cls');
    Add('listexplorers');
    Add('list');
    Add('center');
    Add('cmd');
    Add('hexview');
    Add('preview');
    Add('tojpg');
    EndUpdate;
  end;
end;

procedure TForm1.populateEnvExecutables;
var
  Envs, Env : PChar;
  PathList: TStringList;
  Path: string;
  SR: TSearchRec;
  FilePath: string;
  I: Integer;
begin
  ButtonedEdit1.ACStrings.BeginUpdate;
  ButtonedEdit1.ACStrings.Clear;

  if FEnvExecutables.Count < 1 then
  begin
    Envs := GetEnvironmentStrings;
    PathList := TStringList.Create;
    try
      Env := Envs;
      while Env^ <> #0 do
      begin
        if Pos('PATH=', WideCharToString(Env)) = 1 then
          PathList.DelimitedText := StringReplace(WideCharToString(Env), 'PATH=', '', []);
        Env := Env + StrLen(Env) + 1;
      end;

      for Path in PathList do
      begin
        FilePath := IncludeTrailingPathDelimiter(Path);

        if not DirectoryExists(FilePath) then Continue;

        if FindFirst(FilePath + '*.cmd', faAnyFile, SR) = 0 then
        begin
          repeat
            FEnvExecutables.Add(FilePath + SR.Name);
            ButtonedEdit1.ACStrings.Add(FilePath + SR.Name);
          until FindNext(SR) <> 0;
          FindClose(SR);
        end;
        if FindFirst(FilePath + '*.bat', faAnyFile, SR) = 0 then
        begin
          repeat
            FEnvExecutables.Add(FilePath + SR.Name);
            ButtonedEdit1.ACStrings.Add(FilePath + SR.Name);
          until FindNext(SR) <> 0;
          FindClose(SR);
        end;
        if FindFirst(FilePath + '*.exe', faAnyFile, SR) = 0 then
        begin
          repeat
            FEnvExecutables.Add(FilePath + SR.Name);
            ButtonedEdit1.ACStrings.Add(FilePath + SR.Name);
          until FindNext(SR) <> 0;
          FindClose(SR);
        end;

      end;

    finally
      FreeEnvironmentStringsW(Envs);
      PathList.Free;
    end;
  end
  else
  begin
    for I := 0 to FEnvExecutables.Count - 1 do
    begin
      ButtonedEdit1.ACStrings.Add(FEnvExecutables[I]);
    end;
  end;

  ButtonedEdit1.ACStrings.EndUpdate;

end;

procedure TForm1.populateEnvironmentStrings;
var
  Envs, Env : PChar;
begin
  FCommandType := ctEnvironment;

  ButtonedEdit1.ACStrings.BeginUpdate;
  ButtonedEdit1.ACStrings.Clear;

  Envs := GetEnvironmentStrings;
  try
    Env := Envs;
    while Env^ <> #0 do
    begin
      ButtonedEdit1.ACStrings.Add(Env);
      Env := Env + StrLen(Env) + 1;
    end;
  finally
    FreeEnvironmentStrings(Envs);
  end;

  ButtonedEdit1.ACStrings.EndUpdate;
end;

procedure TForm1.populateMyFolders;
begin
  with ButtonedEdit1.ACStrings do
  begin
    BeginUpdate;
    Clear;
    Add('Dir=L:\Proyectos');
    Add('Dir=F:\Components');
    Add('Dir=L:\FreepascalProjects');
    Add('Dir=F:\Projects');
    Add('Dir=O:\Projects');
    EndUpdate;
  end;
end;

procedure TForm1.ProcessDosCommand(Sender: TObject; ACommand: string; terminateCurrent: Boolean = False);
begin
  if DosCommand1.IsRunning and terminateCurrent then
  begin
    DosCommand1.Stop;
  end
  else if DosCommand1.IsRunning and not terminateCurrent then
  begin
    if MessageDlg('A previous command is processing!'#13#10'Shoul I kill it?', TMsgDlgType.mtWarning, mbYesNo, 0) = mrYes then
    begin
      DosCommand1.Stop;
    end;
  end;

  if not DosCommand1.IsRunning then
  begin
    try
    DosCommand1.InputToOutput := False;

    DosCommand1.CommandLine := ACommand;
    DosCommand1.Execute;
    ActivityIndicator1.Visible := True;
    ActivityIndicator1.Animate := DosCommand1.IsRunning;
    except
      on e:ECreateProcessError do
      begin

      end;
    end;
  end;
end;

procedure TForm1.Show1Click(Sender: TObject);
begin
  Show;
end;

function TForm1.ShowPreview(const FileName: string): Boolean;
var
  wicImg: TWICImage;
  wicList: IWICImagingFactory;
  hr: HRESULT;
  list: IEnumUnknown;
  vInt: IUnknown;
  decoder: IWICBitmapDecoderInfo;
  vBuf: array[0..255] of Char;
  vLen: UINT;
  friendlyName: string;
  fileext: string;
begin

{CoInitialize(nil);

  hr := CoCreateInstance(CLSID_WICImagingFactory, nil, CLSCTX_INPROC_SERVER,
          IID_IWICImagingFactory, wicList);
  //OleCheck(hr);
  if Succeeded(hr) then
  begin
    hr := wicList.CreateComponentEnumerator(WICDecoder, WICComponentEnumerateDefault, list);
    if Succeeded(hr) then
    begin
      while list.Next(1, vInt, nil) = S_OK do
      begin
        if Succeeded(vInt.QueryInterface(IID_IWICBitmapDecoderInfo, decoder)) then
        begin
          if (decoder.GetFriendlyName(High(vBuf), vBuf, vLen) = S_OK) and (vLen > 1) then
          begin
            SetString(friendlyName, PChar(@vBuf), vLen - 1);
            BCEditor1.Lines.Add('WIC : ' + friendlyName);
          end;
          if (decoder.GetFileExtensions(0, nil, vLen) = S_OK) and (vLen > 1) then
          begin
            SetLength(fileext, vLen - 1);
            decoder.GetFileExtensions(vLen, PChar(fileExt), vLen);
            BCEditor1.Lines.Add('WIC extensions: ' + fileext);
          end;

        end;
        vInt := nil;
      end;

    end;
  end;


CoUninitialize;}

  Image1.Picture.Bitmap.Handle := ExtractThumbnail(FileName, 256, 256 );

  Result := False;

  if fPreview <> nil then
    fPreview.Free;
{ DISABLE FOR NOW
  fPreview := THostPreviewHandler.Create(Self);
  fPreview.Top := 0;
  fPreview.Left := 0;
  fPreview.Width := pnlPreview.ClientWidth;
  fPreview.Height := pnlPreview.ClientHeight;
  fPreview.Parent := pnlPreview;
  fPreview.Align := alClient;
  fPreview.FileName := FileName;

  if fPreview.Previewable then
  begin
    fPreview.Visible := True;
    THostPreviewHandlerClass(fPreview).Paint
  end
  else
  begin
    // handle by ourselves the preview
    fPreview.Visible := False;
    case IndexStr(ExtractFileExt(FileName).ToLower,[
    // WIC supported by default are the following according to
    // https://docs.microsoft.com/es-mx/windows/win32/wic/-wic-about-windows-imaging-codec?redirectedfrom=MSDN
      '.bmp', '.gif', '.ico', '.jpeg', '.jpg',
      '.jfif', '.png', '.tiff', '.wdp', '.dds',
    //https://docs.microsoft.com/en-us/windows/win32/wic/native-wic-codecs
      '.dng', '.jxr', '.tif', '.jpe', '.dib',
    // unsupported (unless you installed a wic-enabled codec)
      '.webp', '.avif', '.heif', '.flif'
    ]) of
      0..14:
      begin
        wicImg := TWICImage.Create;
        try

          wicImg.LoadFromFile(FileName);
          EsImage1.Picture.Assign(wicImg);
          EsImage1.Repaint;
        finally
          wicImg.Free;
        end;
      end
      else
      begin

      end;

    end;

  end;}
end;

procedure TForm1.SpeedButton1Click(Sender: TObject);
begin
  with SpeedButton1 do
  begin
    if FPinned then // if not pinned
    begin
      FPinned := False;
      Caption := '';
    end
    else
    begin
      FPinned := True;
      Caption := '' // pin
    end;
  end;

end;

procedure TForm1.SwitchToWindow(AWnd: HWND);
var
  HActiveWindow: HWND;
  HForegroundThread, HAppThread: DWORD;
  FClientId: DWORD;
begin
  HActiveWindow := AWnd;

  HForegroundThread := GetWindowThreadProcessId(HActiveWindow, @FClientId);
  AllowSetForegroundWindow(FClientId);
//  SwitchToThisWindow(AWnd, True);
  HAppThread := GetCurrentThreadId;

  AttachThreadInput(HForegroundThread, HAppThread, True);
  BringWindowToTop(AWnd);
  Winapi.Windows.SetFocus(AWnd);
  AttachThreadInput(HForegroundThread, HAppThread, False);
  SetForegroundWindow(AWnd);
end;

procedure TForm1.tmrOutputTimer(Sender: TObject);
var
  tmpbuf: TStringList;
begin
  if not Assigned(FOutputBuffer) then Exit;

  if FOutputBuffer.Count = 0 then
    Exit;

  tmpbuf := TStringList.Create;
  try
    FSyncLock.Enter;
    try
      tmpbuf.Assign(FOutputBuffer);
      FOutputBuffer.Clear;
    finally
      FSyncLock.Leave;
    end;

    BCEditor1.Lines.BeginUpdate;
    try
      BCEditor1.Lines.AddStrings(tmpbuf);
      BCEditor1.GotoLineAndCenter(BCEditor1.Lines.Count - 1);
      BCEditor1.Refresh;
    finally
      BCEditor1.Lines.EndUpdate;
    end;
  finally
    tmpbuf.Free;
  end;
end;

procedure TForm1.tmrToastTimer(Sender: TObject);
begin
  StatusBar1.Panels[0].Text := CurrentFile;
  tmrToast.Enabled := False;
end;

procedure TForm1.Toast(aText, aTitle, sType: string; ParentBase: TWinControl);
begin
  StatusBar1.Panels[0].Text := aText;
  tmrToast.Enabled := True;

end;

procedure TForm1.TrayIcon1DblClick(Sender: TObject);
begin
  Visible := not Visible;
end;

procedure TForm1.UpdateMainMenu(const ForeGroundWindow: HWND);
var
  lMenu: HMENU;
  MenuItemCount, MenuItemID: Integer;
  MenuItemText: array[0..255] of Char;
  i: Integer;
  NewMenuItem: TMenuItem;

  AU: TCUIAutomation;
  WindowElement, MenuElement, MenuItemElement: IUIAutomationElement;
  Collection, MenuItems: IUIAutomationElementArray;
  Condition: IUIAutomationCondition;
  MenuItemName: WideString;
  Len: Integer;
  retVal: Integer;
  ExpandCollapsePattern: IUIAutomationExpandCollapsePattern;
//const
//  ControlType_Menu: TGUID = '{d9077285-5a2e-4fb1-991c-ac0f69a4d9b3}'; // Menu control type GUID

begin
//  MainMenu1.Items.Clear;

  if ForeGroundWindow = Handle then
    Exit;

  AU := TCUIAutomation.Create(nil);

  AU.ElementFromHandle(Pointer(ForeGroundWindow), WindowElement);

//  AU.CreatePropertyCondition(UIA_ControlTypePropertyId, ControlType_Menu, Condition);
  AU.CreateTrueCondition(Condition);

//  WindowElement.FindFirst(TreeScope_Descendants, Condition, MenuElement);
  WindowElement.FindAll(TreeScope_Descendants, Condition, Collection);

  Collection.Get_Length(Len);

  for I := 0 to Len - 1 do
  begin
    Collection.GetElement(I, MenuItemElement);
    MenuItemElement.Get_CurrentControlType(retVal);

    if (retVal = UIA_MenuItemControlTypeId) then
    begin
      MenuItemElement.Get_CurrentName(MenuItemName);

//      NewMenuItem := TMenuItem.Create(MainMenu1);
//      NewMenuItem.Caption := MenuItemName;
//      NewMenuItem.Tag := I;
      BCEditor1.Lines.Add(MenuItemName);
//      MainMenu1.Items.Add(NewMenuItem);

//      MenuItemElement.GetCurrentPattern(UIA_ExpandCollapsePatternId, IInterface(ExpandCollapsePattern));
//      if Assigned(ExpandCollapsePattern) then
//      begin
//        ExpandCollapsePattern.Expand;
//        if Recurse = True then
//
//
//      end;

    end;

  end;

//  if Assigned(MenuElement) then
//  begin
//    MenuElement.FindAll(TreeScope_Children, Condition, MenuItems);
//    if Assigned(MenuItems) then
//    begin
//      MenuItems.Get_Length(Len);
//      for I := 0 to Len - 1 do
//      begin
//        MenuItems.GetElement(I, MenuItemElement);
//        MenuItemElement.GetCurrentPropertyValue(UIA_NamePropertyId, MenuItemName);
//
//        NewMenuItem := TMenuItem.Create(MainMenu1);
//        NewMenuItem.Caption := MenuItemName;
//        NewMenuItem.Tag := I;
//        BCEditor1.Lines.Add(MenuItemName);
//
//        MainMenu1.Items.Add(NewMenuItem);
//      end;
//
//    end;
//  end;

  AU.Free;

//  lMenu := GetMenu(FindWindow('TAppBuilder', 'ExplorerCommand - RAD Studio 11 - main [Built]'));
//  if lMenu <> 0 then
//  begin
//    MenuItemCount := GetMenuItemCount(lMenu);
//    for I := 0 to MenuItemCount - 1 do
//    begin
//      MenuItemID := GetMenuItemID(lMenu, I);
//      if MenuItemID <> -1 then
//      begin
//        GetMenuString(lMenu, MenuItemID, MenuItemText, SizeOf(MenuItemText), MF_BYCOMMAND);
//        NewMenuItem := TMenuItem.Create(MainMenu1);
//        NewMenuItem.Caption := MenuItemText;
//        NewMenuItem.Tag := MenuItemID;
//        MainMenu1.Items.Add(NewMenuItem);
//      end;
//    end;
//  end;
end;

procedure TForm1.UpdateStyle;
const
  BGCOLOR = $00191919;//$00362A28;
begin
  //on light
  if IsWindowsDarkMode then
  begin
    AllowDarkModeForApp(True);
    Form1.Color := RGB(38, 40, 4);
    Form1.AlphaBlend := True;
    Form1.AlphaBlendValue := 253;
    with SynPasSyn1 do
    begin
      CommentAttri.Foreground := $00A47262;
      CommentAttri.Background := BGCOLOR;

//      EventAttri.Foreground := $00FDE98B;
//      EventAttri.Background := $00362A28;
//      EventAttri.Style := [fsBold];

      IdentifierAttri.Foreground := $00F2F8F8;
      IdentifierAttri.Background := BGCOLOR;

      KeyAttri.Foreground := $0054B91D;//FDE98B;
      KeyAttri.Background := BGCOLOR;
      KeyAttri.Style := [fsBold];

//      NonReservedKeyAttri.Foreground := $0054B91D;//$00FDE98B;
//      NonReservedKeyAttri.Background := $00362A28;
//      NonReservedKeyAttri.Style := [fsBold];

      NumberAttri.Foreground := $00F993BD;
      NumberAttri.Background := BGCOLOR;

      SpaceAttri.Foreground := clWindowText;
      SpaceAttri.Background := BGCOLOR;//MOST PART

//      SpecVarAttri.Foreground := $00C679FF;
//      SpecVarAttri.Background := $00362A28;
//      SpecVarAttri.Style := [fsBold];

      StringAttri.Foreground := $008BE9FC;
      StringAttri.Background := clNone;

      SymbolAttri.Foreground := $00C679FF;
      SymbolAttri.Background := BGCOLOR;

//      TemplateAttri.Foreground := $008BE9FC;
//      TemplateAttri.Background := clNone;
    end;

    rkSmartPath1.Font.Color := clWhite;
    TStyleManager.TrySetStyle('Windows11 Modern Dark');
  end
  else
  begin
    Form1.Color := RGB(248, 249, 253); //dark: 38, 40 44
    Form1.AlphaBlend := True;
    Form1.AlphaBlendValue := 250; // 253
    TStyleManager.TrySetStyle('Windows');
  end;

end;

procedure TForm1.UpdateTheme;
begin
  UpdateStyle;

//  EnableImmersiveDarkMode(True);
//  UseImmersiveDarkMode(Handle, True); //my function to dark mode titlebar win11+
//  EnableNCShadow(Handle);

  if IsWindowsDarkMode then
  begin
    ACLApplicationController1.DarkMode := TACLBoolean.True;
    SetDarkMode(Handle, True);
  end
  else
  begin
    ACLApplicationController1.DarkMode := TACLBoolean.False;
    SetDarkMode(Handle, False);
  end;
end;

procedure TForm1.RefreshEnvironmentVariables;
var
  TokenHandle: THandle;
  EnvironmentStrings: PEnvironment; // LPTSTR;
  Current: PChar;
begin
  TokenHandle := 0;
  try
    if not OpenProcessToken(GetCurrentProcess, TOKEN_QUERY or TOKEN_DUPLICATE, TokenHandle) then
      RaiseLastOSError;

    // Get the environment strings block
    //EnvironmentStrings := GetEnvironmentStrings;
    if CreateEnvironmentBlock(EnvironmentStrings, TokenHandle, False) then
    try
      if EnvironmentStrings = nil then
        Exit;

      FEnvStrings.Clear;
      FEnvStrings.Delimiter := ';';
      FEnvStrings.StrictDelimiter := True;
      // Loop through the environment strings and reload them
      Current := PChar(EnvironmentStrings);
      while Current^ <> #0 do
      begin
        var EnvEntry := String(Current);
        var Pos := EnvEntry.IndexOf('=');
        if Pos > 0 then
        begin
          var Name := Copy(EnvEntry, 1, Pos);
          var Value := Copy(EnvEntry, Pos + 2, Length(EnvEntry) - Pos - 1);
          Winapi.Windows.SetEnvironmentVariable(PChar(Name), PChar(Value));
          if LowerCase(Name) = 'path' then
            FEnvStrings.DelimitedText := PChar(Value);
        end;

        // Move to the next environment string
        Inc(Current, StrLen(Current) + 1);
      end;
    finally
      //FreeEnvironmentStrings(EnvironmentStrings);
      RtlDestroyEnvironment(EnvironmentStrings);
    end
    else
      RaiseLastOSError;
  finally
    if TokenHandle <> 0 then
      CloseHandle(TokenHandle);
  end;
end;

procedure TForm1.WMSettingChange(var Msg: TMessage);
begin
  if PChar(Msg.LParam) = 'Environment' then
  begin
    RefreshEnvironmentVariables;
//    ShowMessage('Environment refreshed!');
  end;
  inherited;
end;

procedure TForm1.WndProc(var Message: TMessage);
begin
  inherited;

  if Message.Msg = WM_SETTINGCHANGE then
  begin
    UpdateTheme;
  end;
end;

{ ThumbThread }

constructor ThumbThread.Create(View: TrkView; Items: TList);
begin
  ViewLink := View;
  ItemsLink := Items;
  FreeOnTerminate := False;
  inherited Create(False);
  Priority := tpLower;
end;

procedure ThumbThread.Execute;
var
  Cnt, I: Integer;
  PThumb: PItemData;
  Old: Integer;
  InView: Integer;
  ShellFolder, DesktopShellFolder: IShellFolder;
  XtractImage: IExtractImage;
  XtractImage2: IExtractImage2;
  XtractIcon: IExtractIcon;
  fileShellItemImage: IShellItemImageFactory;
  ImageFactory: IShellItemImageFactory;
  Bmp: TBitmap;
  Path: string;
  Eaten: DWORD;
  PIDL: PItemIDList;
  RunnableTask: IRunnableTask;
  Flags: DWORD;
  Buf: array[0..MAX_PATH * 4] of WideChar;
  BmpHandle: HBITMAP;
  Attribute, Priority: DWORD;
  GetLocationRes: HRESULT;
  ThumbJPEG: TJPEGImage;
  MS: TMemoryStream;
  ASize: TSize;
  FName: string;
  p, pro: Integer;
  PV: Single;
  IIdx: Integer;
  IFlags: Cardinal;
  SIcon, LIcon: HICON;
  IconS, IconL: TIcon;
  Done: Boolean;
  Res: HRESULT;
  ColorDepth: Cardinal;
  IsVistaOrLater: Boolean;
begin
  inherited;
  if (ViewLink.Items.Count = 0) then
    Exit;

  IsVistaOrLater := CheckWin32Version(6);

  CoInitializeEx(nil, COINIT_APARTMENTTHREADED or COINIT_DISABLE_OLE1DDE);
  try
    ThumbJPEG := TJPEGImage.Create;
    ThumbJPEG.CompressionQuality := 80;
    ThumbJPEG.Performance := jpBestSpeed;
    Path := form1.Directory;

    OleCheck(SHGetDesktopFolder(DesktopShellFolder));
    OleCheck(DesktopShellFolder.ParseDisplayName(0, nil, StringToOleStr(Path),
      Eaten, PIDL, Attribute));
    OleCheck(DesktopShellFolder.BindToObject(PIDL, nil, IID_IShellFolder,
      Pointer(ShellFolder)));
    CoTaskMemFree(PIDL);

    Cnt := 0;
    Old := ViewLink.ViewIdx;
    pro := 0;
    PV := 100 / ViewLink.Items.Count;
    repeat
      while (not Terminated) and (Cnt < ViewLink.Items.Count) do
      begin
        if Old <> ViewLink.ViewIdx then
        begin
          Cnt := ViewLink.ViewIdx - 1;
          if Cnt = -1 then
            Cnt := 0;
          Old := ViewLink.ViewIdx;
        end;

        PThumb := PItemData(ItemsLink.Items[ViewLink.Items[Cnt]]);
        Done := PThumb.GotThumb;
        PThumb.ImgState := 0;

        if IsVistaOrLater then
        begin
          if not Done then
          begin
            Bmp := TBitmap.Create;
            Bmp.Canvas.Lock;
            FName := Path + PThumb.Name;
            Res := SHCreateItemFromParsingName(PChar(FName), nil,
              IShellItemImageFactory, fileShellItemImage);
            if Succeeded(Res) then
            begin
              ASize.cx := 256;
              ASize.cy := 256;
              Res := fileShellItemImage.GetImage(ASize, SIIGBF_THUMBNAILONLY or SIIGBF_BIGGERSIZEOK,
                BmpHandle);
              if Succeeded(Res) then
              begin
                Bmp.Canvas.Unlock;
                Bmp.Handle := BmpHandle;
                Bmp.Canvas.Lock;
                HackAlpha(Bmp, clWhite);
                PThumb.IsIcon := False;
                Done := True;
              end;
            end;
          end;
        end
        else
        begin
          if not Done then
          begin
            Bmp := TBitmap.Create;
            Bmp.Canvas.Lock;
            OleCheck(ShellFolder.ParseDisplayName(0, nil,
              StringToOleStr(PThumb.Name), Eaten, PIDL, Attribute));
            ShellFolder.GetUIObjectOf(0, 1, PIDL, IExtractImage, nil,
              XtractImage);
            CoTaskMemFree(PIDL);
            if Assigned(XtractImage) then
            begin
              if XtractImage.QueryInterface(IID_IExtractImage2,
                Pointer(XtractImage2)) <> E_NOINTERFACE then
              else
                XtractImage2 := nil;
              RunnableTask := nil;
              ASize.cx := 256;
              ASize.cy := 256;
              Priority := 0;
              Flags :=
                IEIFLAG_SCREEN or IEIFLAG_OFFLINE or IEIFLAG_ORIGSIZE
                or IEIFLAG_QUALITY;
              ColorDepth := 32;
              GetLocationRes := XtractImage.GetLocation(Buf, MAX_PATH,
                Priority, ASize, ColorDepth, Flags);
              if (GetLocationRes = NOERROR) or (GetLocationRes = E_PENDING) then
              begin
                if GetLocationRes = E_PENDING then
                  if XtractImage.QueryInterface(IRunnableTask, RunnableTask)
                    <> S_OK then
                    RunnableTask := nil;
                try
                  if Succeeded(XtractImage.Extract(BmpHandle)) then
                  begin
                    Bmp.Canvas.Unlock;
                    Bmp.Handle := BmpHandle;
                    Bmp.Canvas.Lock;
                    HackAlpha(Bmp, clWhite);
                    PThumb.IsIcon := False;
                    Done := True;
                  end;
                except
                  on E: EOleSysError do
                    OutputDebugString(
                      PChar(string(E.ClassName) + ': ' + E.Message)
                    )
                  else
                    raise;
                end;
              end;
            end;
          end;
        end;

      end;
    until (Cnt = 0) or (Terminated);

    if not Terminated then
      PostMessage(Form1.Handle, CM_UpdateView, 0, 0);
    PostMessage(Form1.Handle, CM_Progress, 0, 100);
    ThumbJPEG.Free;
  finally
    CoUninitialize;
  end;

end;

{ TEnumString }

function TEnumString.Clone(out enm: IEnumString): HResult;
begin
  Result := E_NOTIMPL;
  Pointer(enm) := nil;
end;

constructor TEnumString.Create;
begin
  inherited Create;
  FStrings := TStringList.Create;
  FCurrIndex := 0;
end;

destructor TEnumString.Destroy;
begin
  FStrings.Free;
  inherited;
end;

function TEnumString.Next(celt: Longint; out elt;
  pceltFetched: PLongint): HResult;
var
  I: Integer;
  wStr: WideString;
begin
  I := 0;
  while (I < celt) and (FCurrIndex < FStrings.Count) do
  begin
    wStr := FStrings[FCurrIndex];
    TPointerList(elt)[I] := CoTaskMemAlloc(2 * (Length(wStr) + 1));
    StringToWideChar(wStr, TPointerList(elt)[I], 2 * (Length(wStr) + 1));
    Inc(I);
    Inc(FCurrIndex);
  end;
  if pceltFetched <> nil then
    pceltFetched^ := I;
  if I = celt then
    Result := S_OK
  else
    Result := S_FALSE;
end;

function TEnumString.Reset: HResult;
begin
  FCurrIndex := 0;
  Result := S_OK;
end;

function TEnumString.Skip(celt: Longint): HResult;
begin
  if (FCurrIndex + celt) <= FStrings.Count then
  begin
    Inc(FCurrIndex, celt);
    Result := S_OK;
  end
  else
  begin
    FCurrIndex := FStrings.Count;
    Result := S_FALSE;
  end;
end;

{ TButtonedEdit }

constructor TButtonedEdit.Create(AOwner: TComponent);
begin
  inherited;
  FACList := TEnumString.Create;
  FEnumString  := FACList;
  FACEnabled := True;
  FACOptions := [acAutoSuggest, acUpDownKeyDropsList];
end;

class constructor TButtonedEdit.Create;
begin
  if not TStyleManager.IsCustomStyleActive then
  begin
    Winapi.Windows.Beep(400, 1000);
    TCustomStyleEngine.UnRegisterSysStyleHook('SysListView32', TSysListViewStyleHook);
    TCustomStyleEngine.RegisterSysStyleHook('SysListView32', TSysListViewStyleHook);
  end;
end;

procedure TButtonedEdit.CreateWnd;
var
  Dummy: IUnknown;
  Strings: IEnumString;
  FuzzyMatchList: TStringList;
  FuzzyMatcher: TFuzzyStringMatcher;
  AutocompleteEx: IAutoComplete2;
begin
  inherited;
//  SetWindowTheme(Handle, PChar('DarkMode_Explorer'), nil);
  if HandleAllocated then
  begin
    try
      Dummy := CreateComObject(CLSID_AutoComplete);
      if (Dummy <> nil) and
        (Dummy.QueryInterface(IID_IAutoComplete, FAutoComplete) = S_OK) then
      begin
        //https://learn.microsoft.com/en-us/windows/win32/api/shldisp/ne-shldisp-autocompleteoptions
        // set auto completion options
        if Dummy.QueryInterface(IID_IAutoComplete2, AutoCompleteEx) = S_OK then
          AutoCompleteEx.SetOptions(ACO_AUTOSUGGEST or ACO_AUTOAPPEND or ACO_UPDOWNKEYDROPSLIST);

        case FACSource of
//          acsList: ;
          //It is used to manage the history of autocomplete entries.
          acsHistory: Strings := CreateComObject(CLSID_ACLHistory) as IEnumString;
          //It is used to manage the MRU autocomplete entries.
          acsMRU: Strings := CreateComObject(CLSID_ACLMRU) as IEnumString;
          //It is used to manage autocomplete entries specific to shell folders.
          acsShell:
          begin
            Strings := CreateComObject(CLSID_ACListISF) as IEnumString;
          end
          else
          begin
            // Use FuzzyStringMatch to perform fuzzy string matching
            FuzzyMatchList := TStringList.Create;
            try
              FuzzyMatcher := TFuzzyStringMatcher.Create(8);
            finally

            end;
            Strings := FACList as IEnumString; // original
          end;
        end;
        if S_OK = FAutoComplete.Init(Handle, Strings, nil, nil) then
        begin
          SetACEnabled(FACEnabled);
          SetACOptions(FACOptions);
//          TCustomStyleEngine.RegisterSysStyleHook('SysListView32', TSysListViewStyleHook);
//          TCustomStyleEngine.RegisterSysStyleHook('SysListView32', TSysListViewStyleHook);
        end;
      end;
    except
      // CLSID_IAutoComplete is not available
    end;
  end;

end;

destructor TButtonedEdit.Destroy;
begin
  FACList := nil;
  inherited;
end;

procedure TButtonedEdit.DestroyWnd;
begin
  if (FAutoComplete <> nil) then
  begin
    FAutoComplete.Enable(False);
    FAutoComplete := nil;
  end;

  inherited;

end;

function TButtonedEdit.GetACStrings: TStringList;
begin
  Result := FACList.FStrings;
end;

procedure TButtonedEdit.SetACEnabled(const Value: Boolean);
begin
  if (FAutoComplete <> nil) then
  begin
    FAutoComplete.Enable(FACEnabled);
  end;
  FACEnabled := Value;
end;

procedure TButtonedEdit.SetACOptions(const Value: TACOptions);
const
  Options : array[TACOption]
    of Integer = (
      ACO_NONE,
      ACO_AUTOSUGGEST,
      ACO_AUTOAPPEND,
      ACO_SEARCH,
      ACO_FILTERPREFIXES,
      ACO_USETAB,
      ACO_UPDOWNKEYDROPSLIST,
      ACO_RTLREADING,
      ACO_WORD_FILTER,
      ACO_NOPREFIXFILTERING
      );
var
  Option: TACOption;
  Opt: DWORD;
  AC2: IAutoComplete2;
begin
  if (FAutoComplete <> nil) then
  begin
    if S_OK = FAutoComplete.QueryInterface(IID_IAutoComplete2, AC2) then
    begin
      Opt := ACO_NONE;
      for Option := Low(Options) to High(Options) do
      begin
        if (Option in FACOptions) then
          Opt := Opt or DWORD(Options[Option]);
      end;
      AC2.SetOptions(Opt);
    end;
  end;
  FACOptions := Value;
end;

procedure TButtonedEdit.SetACSource(const Value: TACSource);
begin
  if FACSource <> Value then
  begin
    FACSource := Value;
    RecreateWnd;
  end;
end;

procedure TButtonedEdit.SetACStrings(const Value: TStringList);
begin
  if Value <> FACList.FStrings then
    FACList.FStrings.Assign(Value);
end;

{ TFuzzyStringMatcher }

constructor TFuzzyStringMatcher.Create(Threshold: Integer);
begin
  FThreshold := Threshold;
end;

function TFuzzyStringMatcher.DamerauLevenshteinDistance(const S1,
  S2: string): Integer;
var
  Len1, Len2, I, J, Cost, PrevCost: Integer;
  D: array of array of Integer;
begin
  Len1 := Length(S1);
  Len2 := Length(S2);
  SetLength(D, Len1 + 1, Len2 + 1);;

  for I := 0 to Len1 do
    D[I, 0] := I;

  for J := 0 to Len2 do
    D[0, J] := J;

  for I := 1 to Len1 do
  begin
    for J := 1 to Len2 do
    begin
      if S1[I] = S2[J] then
        Cost := 0
      else
        Cost := 1;

      PrevCost := D[I - 1, J - 1];

      if (I > 1) and (J > 1) and (S1[I - 1] = S2[J]) and (S1[I]  = S2[J - 1]) then
        PrevCost := Min(PrevCost, D[I - 2, J - 2]);

      D[I, J] := Min(Min(D[I - 1, J] + 1, D[I, J - 1] + 1), PrevCost + Cost);
    end;
  end;

  Result := D[Len1, Len2];
end;

function TFuzzyStringMatcher.IsMatch(const Str, SubStr: string): Boolean;
begin
  Result := DamerauLevenshteinDistance(Str, SubStr) <= FThreshold;
end;

end.
