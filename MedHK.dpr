program MedHK;

uses
  Vcl.Forms,
  main in 'main.pas' {Form1},
  datamodul in 'datamodul.pas' {dm: TDataModule};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TForm1, Form1);
  Application.CreateForm(Tdm, dm);
  Application.Run;
end.
