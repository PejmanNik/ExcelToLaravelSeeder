program Excel2LaravelSeeder;

uses
  Vcl.Forms,
  Frm_MainU in 'Frm_MainU.pas' {Form1},
  Vcl.Themes,
  Vcl.Styles;

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  TStyleManager.TrySetStyle('Glow');
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
