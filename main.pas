unit main;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, SynEdit, DosCommand, rkShellPath, rkEdit, UCL.Form,
  Vcl.StdCtrls, Vcl.ExtCtrls, cyButtonedEdit, System.ImageList, Vcl.ImgList,
  BCEditor.Highlighter, BCEditor.Editor, Vcl.ComCtrls, Vcl.WinXCtrls,
  VirtualTrees, TlHelp32, ShellApi, ShDocVw, ActiveX, ShlObj, IniFiles;

const
  KeyEvent = WM_USER + 1;
type
  TForm1 = class(TUForm)
    DosCommand1: TDosCommand;
    ButtonedEdit1: TButtonedEdit;
    ImageList1: TImageList;
    BCEditor1: TBCEditor;
    StatusBar1: TStatusBar;
    SearchBox1: TSearchBox;
    Timer1: TTimer;
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
    procedure Timer1Timer(Sender: TObject);
  private
    { Private declarations }
    lastExplorerHandle: HWND;
    lstExplorer: TStringList;
    lstExplorerWnd: TStringList;
    lstExplorerItem: TStringList;

    function ListExplorerInstances:Integer;
    procedure KeyEventHandler(var Msg: TMessage); message KeyEvent;
    procedure OnFocusLost(Sender: TObject);
  protected
    procedure CreateParams(var Params: TCreateParams); override;
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

  function StartHook:BOOL; stdcall; external 'HotkeyHook.dll' name 'STARTHOOK';
  procedure StopHook; stdcall; external 'HotkeyHook.dll' name 'STOPHOOK';
  procedure SwitchToThisWindow(h1: hWnd; x: bool); stdcall;
  external user32 Name 'SwitchToThisWindow';

implementation

{$R *.dfm}

procedure TForm1.ButtonedEdit1Enter(Sender: TObject);
begin
 ButtonedEdit1.RightButton.Visible := True;
end;

procedure TForm1.ButtonedEdit1KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var
  I: Integer;
begin
  if key = 13 then
  begin
    if ButtonedEdit1.Text = 'list' then
    begin
      ListExplorerInstances;
      BCEditor1.Text := 'Current HWND: ' + IntToStr(lastExplorerHandle) + '';
      for I := 0 to lstExplorer.Count - 1 do
      begin
        if lstExplorerWnd[I] = IntToStr(lastExplorerHandle)  then

        BCEditor1.Text := BCEditor1.Text + #13#10 + lstExplorer[I] + ' ' + lstExplorerWnd[i];
      end;
    end
    else if ButtonedEdit1.Text = 'items' then
    begin
      ListExplorerInstances;
      BCEditor1.Text := 'Current HWND: ' + IntToStr(lastExplorerHandle) + '';
      for I := 0 to lstExplorerItem.Count - 1 do
      begin
        BCEditor1.Text := BCEditor1.Text + #13#10 + lstExplorerItem[I];
      end;

    end
    else if ButtonedEdit1.Text = 'exit' then
      close
    else
    begin

    try
      BCEditor1.Lines.Clear;
      DosCommand1.CurrentDir := ExtractFilePath(ParamStr(0));
      DosCommand1.CommandLine := 'cmd.exe /c ' + ButtonedEdit1.Text;
      DosCommand1.Execute;
    except
      //on E:Exception do

    end;

    end;
    ButtonedEdit1.Text := '';
 end;
end;

procedure TForm1.CreateParams(var Params: TCreateParams);
begin
  inherited;

  Params.WinClassName := 'ExplorerCommandWnd';
end;

procedure TForm1.DosCommand1ExecuteError(ASender: TObject; AE: Exception;
  var AHandled: Boolean);
begin
  BCEditor1.Lines.Text := AE.ToString;
end;

procedure TForm1.DosCommand1NewLine(ASender: TObject; const ANewLine: string;
  AOutputType: TOutputType);
begin
//  BCEditor1.Lines.Add(ANewLine);
  BCEditor1.Text := BCEditor1.Text +#13#10+ ANewLine;
  BCEditor1.Perform(EM_SCROLL, SB_LINEDOWN, 0);
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  if not StartHook then
  begin
    MessageDlg('Couldn''t set global hotkey.',mtError, [mbOK], 0);
    Application.Terminate;
  end;

  KeyPreview := True;

  lstExplorer := TStringList.Create;
  lstExplorerWnd := TStringList.Create;
  lstExplorerItem := TStringList.Create;

  Application.OnDeactivate := OnFocusLost;
end;

procedure TForm1.FormDestroy(Sender: TObject);
begin
  lstExplorerWnd.Free;
  lstExplorer.Free;
  lstExplorerItem.Free;

  StopHook;
end;

procedure TForm1.FormKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if key = VK_ESCAPE then
    Hide;
end;

procedure TForm1.FormShow(Sender: TObject);
begin
  ShowWindow(Application.Handle, SW_HIDE);
end;

procedure TForm1.KeyEventHandler(var Msg: TMessage);
var
  I: Integer;
  command: String;
  Ret: Integer;
begin
  command := PChar(Msg.LParam);
  lastExplorerHandle := StrToInt(command);

  if not Visible then
  begin
    Show;
    SetForegroundWindow(Handle);
    //SwitchToThisWindow(Handle, True);

    Ret := ListExplorerInstances;

    StatusBar1.Panels[0].Text := IntToStr(Ret);

    for I := 0 to lstExplorerWnd.Count - 1 do
    begin
      if lstExplorerWnd[i] = IntToStr(lastExplorerHandle) then
        StatusBar1.Panels[0].Text := lstExplorerItem[i];
    end;
  end;

end;

// Lists explorer instances which has items visible, ignores special directories
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
begin
  Result := 0;
  lstExplorer.BeginUpdate;
  lstExplorer.Clear;
  lstExplorerWnd.BeginUpdate;
  lstExplorerWnd.Clear;
  lstExplorerItem.BeginUpdate;
  lstExplorerItem.Clear;

  hr := CoInitialize(nil); // <-- manually call CoInitialize()
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
                          folderPath := StrAlloc(MAX_PATH);
                          if SHGetPathFromIDList(ItemIDList2, folderPath) then
                            lstExplorer.Add(folderPath);
                          SHGetMalloc(AMalloc);
                          AMalloc.Free(ItemIDList2);
                          StrDispose(folderPath);
                        end;

                        if Succeeded(PersistFolder2.QueryInterface(IID_IShellFolder, ShellFolder)) then
                        begin
                          if Succeeded(ShellFolder.GetDisplayNameOf(ItemIDList, SHGDN_FORPARSING, ret)) then
                            lstExplorerItem.Add(ret.pOleStr);
                        end;
                      end;
                    end;
                  end;

                end;

              end;
            end;
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
  lstExplorer.EndUpdate;
end;

procedure TForm1.OnFocusLost(Sender: TObject);
begin
  StatusBar1.Panels[0].Text := '';
  Hide;
end;

procedure TForm1.Timer1Timer(Sender: TObject);
begin
{TODO: improve detection for other reasons, like keeping it even though it changes to other explorer instance}
  if (GetForegroundWindow <> Self.Handle)
  then
  begin
    lastExplorerHandle := GetForegroundWindow;
  end;
end;

end.
