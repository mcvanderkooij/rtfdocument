unit UfrmTest;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls;

type
  TForm1 = class(TForm)
    Button1: TButton;
    Memo1: TMemo;
    RichEdit1: TRichEdit;
    Label1: TLabel;
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

uses
  URTFDocument;

{$R *.dfm}

procedure TForm1.Button1Click(Sender: TObject);
var
  LDocument: TRtfDocument;
  LStream: TMemoryStream;

  par: TRtfParagraph;
  table: TRtfTable;

  LColorIndexRed: TRtfColorDescriptor;
  LColorIndexBlue: TRtfColorDescriptor;
  LColorIndexYellow: TRtfColorDescriptor;

  LFontIndexMonotypeCorsiva: TRtfFontDescriptor;
  LFontIndexTahoma: TRtfFontDescriptor;

  LList: TRtfList;
  LListItem: TRtfListItem;
  LListItems: TRtfListItems;
begin
  LStream := TMemoryStream.Create;
  LDocument := TRtfDocument.Create(psA4, poPortrait, lcEnglish);
  try
    LColorIndexRed := LDocument.CreateColor(clRed);
    LColorIndexBlue := LDocument.CreateColor(clBlue);
    LColorIndexYellow := LDocument.CreateColor(clYellow);

    LFontIndexMonotypeCorsiva := LDocument.CreateFont('Monotype Corsiva');
    LFontIndexTahoma := LDocument.CreateFont('Tahoma');

    LDocument.DefaultCharFormat.Font := LFontIndexTahoma;
    LDocument.DefaultCharFormat.FontSize := 8;

    par := LDocument.AddParagraph;
    par.Text.AppendLine('Salvête, Omnês!');
    par.DefaultCharFormat.FontSize := 60;
    par.DefaultCharFormat.Font := LFontIndexMonotypeCorsiva;
    par.DefaultCharFormat.FontStyle.StyleAdd := [fsBold, fsItalic];
    par.Alignment := alCenter;

    par := LDocument.AddParagraph;
    par.Text.Append('Some test paragraph');
    par.AddCharFormat(6,9).FontStyle.AddStyle(fsBold);
    par.AddCharFormat(8,12).FontStyle.AddStyle(fsItalic);
    par.AddCharFormat(13,14).FGColor := LColorIndexRed;

    LDocument.AddSpecialInsert(TRtfSpecialInsertHR.Create);

    par := LDocument.AddParagraph;
    par.Text.Append('Paragraph 2');

    par.Text.AppendLine('Unicode: ܔ');

    table := LDocument.AddTable(4, 5, 12);
    table.HeaderBackgroundColour := LColorIndexRed;
    table.RowBackgroundColour := LColorIndexBlue;
    table.RowAltBackgroundColour := LColorIndexYellow;
    par := table.Cell[0, 0].AddParagraph;
    par.Text.Append('row0col0');

    table.Merge(1, 2, 2, 3);

    table.SetInnerBorder(bsNone, 2);
    table.SetOuterBorder(bsDouble, 2);
    table.Cell[3, 1].Borders[drTop].Style := bsDotted;

    LDocument.AddImage('test1.png').Alignment := alRight;
    LDocument.AddParagraph;
//    LDocument.AddImage('test2.jpg');
    LDocument.AddParagraph;
    LDocument.AddImage('test3.bmp').Alignment := alCenter;
    LDocument.AddParagraph;
    LDocument.AddImage('test2.jpg').Alignment := alCenter;
    LDocument.AddParagraph;

    par := LDocument.AddParagraph;
    par.SetText('This is a text with a footnote');
    par := par.AddFootnote(3).AddParagraph;
    par.SetText('This is the footnote');

    LDocument.AddParagraph;
    LDocument.AddParagraph;

    LList := LDocument.AddList(ltRomanUpper, LDocument);
    LList.StartNumber := 18;
    LList.DefaultCharFormat.FGColor := LColorIndexBlue;
    LList.DefaultCharFormat.BGColor := LColorIndexRed;
    LListItem := LList.AddItem;
    LListItem.Paragraph.SetText('Designates paragraph style. If a paragraph style is specified, style properties must be specified with the paragraph. N references an entry in the stylesheet.');

    LListItem := LList.AddItem;
    LListItem.Paragraph.SetText('Toggles automatic hyphenation for the \paragraph. Append 1 or nothing to toggle property on; append 0 to turn it off.');


    LListItems := LList.AddList(ltArrow);
    LListItem := LListItems.AddItem;
    LListItem.Paragraph.SetText('Outline level of paragraph. The N argument is a value from 0-8 representing the outline level of the paragraph. In the default case, no outline level is specified (same as body text).');


//    par := LDocument.Header.AddParagraph;
//    par.SetText('This is a header');
//    par := LDocument.Header.AddParagraph;
//    par.SetText(' van ');
//    par.Alignment := alRight;
//    par.AddControlWord(1, ftPage);
//    par.AddControlWord(4, ftNumPages);
//    LDocument.Header.AddSpecialInsert(TRtfSpecialInsertHR.Create);
//
//    LDocument.Footer.AddSpecialInsert(TRtfSpecialInsertHR.Create);
//    par := LDocument.Footer.AddParagraph;
//    par.SetText('This is a footer');

    memo1.Lines.Text := LDocument.Render;
    RichEdit1.PlainText := False;
    LDocument.ToStream(LStream);
    LStream.Position := 0;
    RichEdit1.Lines.LoadFromStream(LStream);

    LStream.Position := 0;
    LStream.SaveToFile('test.rtf');
  finally
    LDocument.Free;
    LStream.Free;
  end;
end;

end.
