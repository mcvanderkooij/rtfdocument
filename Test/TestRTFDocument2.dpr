program TestRTFDocument2;

uses
  Forms,
  UfrmTest in 'UfrmTest.pas' {Form1},
  URTFDocument in '..\URTFDocument.pas',
  USimpleLinkedList in '..\USimpleLinkedList.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
