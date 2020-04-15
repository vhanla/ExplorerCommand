unit main;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, SynEdit, DosCommand, rkShellPath, rkEdit, UCL.Form,
  Vcl.StdCtrls, Vcl.ExtCtrls, cyButtonedEdit, System.ImageList, Vcl.ImgList,
  BCEditor.Highlighter, BCEditor.Editor, Vcl.ComCtrls;

const
  KeyEvent = WM_USER + 1;
type
  TForm1 = class(TUForm)
    DosCommand1: TDosCommand;
    ButtonedEdit1: TButtonedEdit;
    ImageList1: TImageList;
    BCEditor1: TBCEditor;
    StatusBar1: TStatusBar;
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
  private
    { Private declarations }
    procedure KeyEventHandler(var Msg: TMessage); message KeyEvent;
  protected
    procedure CreateParams(var Params: TCreateParams); override;
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

  procedure StartHook; stdcall; external 'HotkeyHook.dll' name 'STARTHOOK';
  procedure StopHook; stdcall; external 'HotkeyHook.dll' name 'STOPHOOK';

implementation

{$R *.dfm}

procedure TForm1.ButtonedEdit1Enter(Sender: TObject);
begin
 ButtonedEdit1.RightButton.Visible := True;
end;

procedure TForm1.ButtonedEdit1KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if key = 13 then
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
  StartHook;
end;

procedure TForm1.FormDestroy(Sender: TObject);
begin
  StopHook;
end;

procedure TForm1.FormShow(Sender: TObject);
begin
  ShowWindow(Application.Handle, SW_HIDE);
end;

procedure TForm1.KeyEventHandler(var Msg: TMessage);
begin

end;

end.
