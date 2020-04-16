program ExplorerCommand;

uses
  Vcl.Forms,
  Windows,
  SysUtils,
  main in 'main.pas' {Form1},
  Vcl.Themes,
  Vcl.Styles;

{$R *.res}

begin
  if CreateMutex(nil, True, '{C97C27E2-C5FC-41BE-AF34-6C9E250FC303}') = 0 then
    RaiseLastOSError;
  if GetLastError = ERROR_ALREADY_EXISTS then
    Exit;

  Application.Initialize;
  Application.MainFormOnTaskbar := False;
  Application.ShowMainForm := False;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
