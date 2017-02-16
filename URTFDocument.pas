unit URTFDocument;

interface

{ Translation of the c# https://github.com/elistia/DotNetRtfWriter
  I did ignore most of the remarks and comments.
}

uses
  Windows, Classes, SysUtils, Graphics, Generics.Collections, USimpleLinkedList;

type
  TRtfPaperSize = (psA5, psA4, psA3, psLetter, psLegal);
  TRtfPaperOrientation = (poPortrait, poLandscape);
  TRtfLcid = (lcTraditionalChinese = 1028,
        lcEnglish = 1033,
        lcFrench = 1036,
        lcGerman = 1031,
        lcItalian = 1040,
        lcJapanese = 1041,
        lcKorean = 1042,
        lcSimplifiedChinese = 2052,
        lcSpanish = 3082);
  TRtfAlign = (alNone, alLeft, alRight, alCenter, alFullyJustified, alDistributed);
  TRtfAlignVertical = (avTop, avMiddle, avBottom);
  TRtfDirection = (drTop, drRight, drBottom, drLeft);
  TRtfFontStyleFlag = (fsBold, fsItalic, fsUnderline, fsSuper, fsSub, fsScaps, fsStrike);
  TRtfFontStyleFlags = set of TRtfFontStyleFlag;
  TRtfImageFileType = (itJpg, itGif, itPng, itBmp);
  TRtfBorderStyle = (bsNone, bsSingle, bsDotted, bsDashed, bsDouble);
  TRtfTwoInOneStyle = (toNotEnabled, toNone, toParentheses, toSquareBrackets, toAngledBrackets, toBraces);
  TRtfHeaderFooterType = (hfHeader, hfFooter);
  TRtfSectionStartsEnd = (sseStart, sseEnd);
  TRtfListType = (ltBullet, ltArrow, ltCircle, ltSquare, ltDiamond, ltAlphaUpper,
    ltAlphaLower, ltRomanUpper, ltRomanLower, ltArabicNum);

  TRtfException = Exception;

  TRtfTableCell = class;
  TRtfCharFormat = class;
  TRtfTable = class;
  TRtfSpecialInsert = class;
  TRtfList = class;
  TRtfParagraph = class;
  TRtfImage = class;
  TRtfSection = class;
  TRtfSectionFooter = class;
  TRtfDocument = class;

  TRtfFontStyle = class
  private
    FStyleAdd: TRtfFontStyleFlags;
    FStyleRemove: TRtfFontStyleFlags;
  protected
    constructor Create; overload; virtual;
    constructor Create(AStyle: TRtfFontStyle); overload; virtual;

    procedure CopyFrom(ASource: TRtfFontStyle);
  public
    procedure AddStyle(AStyle: TRtfFontStyleFlag);
    procedure RemoveStyle(AStyle: TRtfFontStyleFlag);
    function IsEmpty: boolean;

    property StyleAdd: TRtfFontStyleFlags read FStyleAdd write FStyleAdd;
    property StyleRemove: TRtfFontStyleFlags read FStyleRemove write FStyleRemove;
  end;

  TRtfDescriptor = integer;
  TRtfFontDescriptor = TRtfDescriptor;
  TRtfColorDescriptor = TRtfDescriptor;

  TRtfMargins = class(TObject)
  private
    FMargins: array[TRtfDirection] of Double;
    function GetMargins(index: TRtfDirection): double;
    procedure SetMargins(index: TRtfDirection; const Value: double);
  protected
    constructor Create; overload; virtual;
    constructor Create(ATop, ARight, ABottom, ALeft: double); overload; virtual;
    procedure CopyFrom(ASource: TRtfMargins);
  public
    function Equals(Obj: TObject): Boolean; override;
    property Margins[index: TRtfDirection]: double read GetMargins write SetMargins; default;
  end;

  TRtfMarginsList = TObjectList<TRtfMargins>;

  TRtfBorder = class
  private
    FStyle: TRtfBorderStyle;
    FWidth: double;
    FColorDesc: TRtfColorDescriptor;
  protected
    constructor Create(); virtual;
  public
    function Equals(Obj: TObject): boolean; override;
    function GetHashCode: integer; override;

    property Style: TRtfBorderStyle read FStyle write FStyle;
    property Width: double read FWidth write FWidth;
    property ColorDesc: TRtfColorDescriptor read FColorDesc write FColorDesc;
  end;

  TRtfBorders = class
  private
    FBorders: array[TRtfDirection] of TRtfBorder;
    function GetBorders(index: TRtfDirection): TRtfBorder;
  protected
    constructor Create(); virtual;
  public
    destructor Destroy; override;
    property Borders[index: TRtfDirection]: TRtfBorder read GetBorders; default;
  end;

  TRtfColor = class
  private
    FColor: integer;
    function GetBlue: string;
    function GetGreen: string;
    function GetRed: string;
  public
    constructor Create; overload; virtual;
    constructor Create(ARed, AGreen, ABlue: byte); overload; virtual;
    constructor Create(AHex: string); overload; virtual;
    constructor Create(AColor: TColor); overload; virtual;

    function Equals(Obj: TObject): boolean; override;
    function GetHashCode: integer; override;

    property Color: integer read FColor;
    property Red: string read GetRed;
    property Green: string read GetGreen;
    property Blue: string read GetBlue;
  end;

  TRtfFontTable = TStringList;
  TRtfColorTable = class(TObjectList<TRtfColor>)
  public
    function IndexOfColor(var AColor: TRtfColor): TRtfColorDescriptor;
  end;

  TRtfCellMergeInfo = class
  private
    FRowSpan: integer;
    FColSpan: integer;
    FRowIndex: integer;
    FColIndex: integer;
    FRepresentative: TRtfTableCell;
  protected
    constructor Create(ARepresentative: TRtfTableCell; ARowSpan, AColSpan,
      ARowIndex, AColIndex: integer); virtual;

    property RowSpan: integer read FRowSpan;
    property ColSpan: integer read FColSpan;
    property RowIndex: integer read FRowIndex;
    property ColIndex: integer read FColIndex;
    property Representative: TRtfTableCell read FRepresentative;
  end;

  TRtfCharFormat = class
  private
    FScopeBegin: integer;
    FScopeEnd: integer;
    FFont: TRtfFontDescriptor;
    FAnsiFont: TRtfFontDescriptor;
    FFontSize: integer;
    FFontStyle: TRtfFontStyle;
    FBGColor: TRtfColorDescriptor;
    FFGColor: TRtfColorDescriptor;
    FTwoInOneStyle: TRtfTwoInOneStyle;
    FBookmark: string;
    FLocalHyperlink: string;
    FLocalHyperlinkTip: string;
  protected
    procedure CopyFrom(ASource: TRtfCharFormat);
    procedure SetRange(AScopeBegin, AScopeEnd, ATextLength: integer);

    function RenderHead: string;
    function RenderTail: string;

    property ScopeBegin: Integer read FScopeBegin;
    property ScopeEnd: Integer read FScopeEnd;
  public
    constructor Create(AScopeBegin, AScopeEnd, ATextLength: integer); virtual;
    destructor Destroy; override;

    procedure SetFont(ARtfDocument: TRtfDocument; AFont: TFont);

    property Bookmark: string read FBookmark write FBookmark;
    property LocalHyperlink: string read FLocalHyperlink write FLocalHyperlink;
    property LocalHyperlinkTip: string read FLocalHyperlinkTip write FLocalHyperlinkTip;

    property Font: TRtfFontDescriptor read FFont write FFont;
    property AnsiFont: TRtfFontDescriptor read FAnsiFont write FAnsiFont;
    property FontSize: integer read FFontSize write FFontSize;
    property FontStyle: TRtfFontStyle read FFontStyle;
    property BGColor: TRtfColorDescriptor read FBGColor write FBGColor;
    property FGColor: TRtfColorDescriptor read FFGColor write FFGColor;
    property TwoInOneStyle: TRtfTwoInOneStyle read FTwoInOneStyle write FTwoInOneStyle;
  end;
  TRtfCharFormats = TObjectList<TRtfCharFormat>;

  TRtfRenderable = class
  public
    function Render: string; virtual; abstract;
  end;

  TRtfBlock = class(TRtfRenderable)
  protected
    FAlignment: TRtfAlign;
    FMargins: TRtfMargins;
    FBlockTail: string;
    FDefaultCharFormat: TRtfCharFormat;
    FStartNewPage: boolean;
    FBlockHead: string;
    function GetDefaultCharFormat: TRtfCharFormat; virtual;
  public
    constructor Create; virtual;
    destructor Destroy; override;

    function AlignmentCode: string;

    property Alignment: TRtfAlign read FAlignment write FAlignment;
    property Margins: TRtfMargins read FMargins;
    property DefaultCharFormat: TRtfCharFormat read GetDefaultCharFormat;
    property StartNewPage: boolean read FStartNewPage write FStartNewPage;
    property BlockHead: string write FBlockHead;
    property BlockTail: string write FBlockTail;

  end;

  TRtfBlocklist = class(TRtfRenderable)
  private
    FBlocks: TObjectList<TRtfBlock>;
    FDefaultCharFormat: TRtfCharFormat;

    FAllowParagraph: Boolean;
    FAllowFootnote: Boolean;
    FAllowControlWord: Boolean;
    FAllowImage: Boolean;
    FAllowTable: Boolean;
    FAllowSpecialInsert: Boolean;
    FAllowList: Boolean;
    function GetDefaultCharFormat: TRtfCharFormat;
  protected
    constructor Create; overload;
    constructor Create(AAllowParagraph, AAllowTable: Boolean); overload;
    constructor Create(AAllowParagraph, AAllowFootnote, AAllowControlWord,
      AAllowImage, AAllowTable, AAllowSpecialInsert, AAllowList: Boolean); overload;
    procedure AddBlock(ABlock: TRtfBlock);
    procedure TransferBlocksTo(ATarget: TRtfBlocklist);
  public
    destructor Destroy; override;
    function Render: string; override;
    function AddParagraph: TRtfParagraph;
    procedure AddSpecialInsert(AInsert: TRtfSpecialInsert);
    function AddList(AType: TRtfListType; ADoc: TRtfDocument): TRtfList;
    function AddSection(AType: TRtfSectionStartsEnd; ADoc: TRtfDocument): TRtfSection;
    function AddImage(AImageFilename: string): TRtfImage; overload;
    function AddImage(AImageStream: TStream): TRtfImage; overload;
    function AddImage(AImageGraphic: TGraphic): TRtfImage; overload;
    function AddTable(ARowCount, AColCount: Integer; AHorizontalWidth, AFontSize: Double): TRtfTable;

    property DefaultCharFormat: TRtfCharFormat read GetDefaultCharFormat;
  end;

  TRtfFootNote = class(TRtfBlocklist)
  private
    FPosition: integer;
  protected
    constructor Create(APosition, ATextLength: integer);

    property Position: integer read FPosition;
  public
    function Render: string; override;
  end;

  TRtfFootNotes = TObjectList<TRtfFootNote>;

  TRtfFieldControlWord = class(TRtfRenderable)
  public
    type
      TFieldType = (ftNone, ftPage, ftNumPages, ftDate, ftTime);
  private
    FPosition: integer;
    FType: TFieldType;
  protected
    constructor Create(APosition: integer; AType: TFieldType);

    property Position: integer read FPosition;
  public
    function Render: string; override;
  end;

  TRtfFieldControlWords = TObjectList<TRtfFieldControlWord>;

  TRtfSpecialInsert = class(TRtfBlock)
  protected
    function Special: string; virtual;
  public
    constructor Create; override;
    function Render: string; override;
  end;

  TRtfSpecialInsertHR = class(TRtfSpecialInsert)
  protected
    function Special: string; override;
  end;



  TRtfListItem = class(TRtfBlock)
  private
    FParagraph: TRtfParagraph;
  protected
    function GetParagraph: TRtfParagraph; virtual;
  public
    destructor Destroy; override;
    function Render: string; override;

    property Paragraph: TRtfParagraph read GetParagraph;
  end;

  TRtfListItems = class(TRtfListItem)
  private
    FParent: TRtfListItems;
    FDocument: TRtfDocument;
    FTextindent: Double;
    FListIndent: Double;
    FItems: TObjectList<TRtfListItem>;
    FListType: TRtfListType;
    FNumberSeparator: string;
    FNumberPrefix: string;
    FNumberSuffix: string;
    FDefaultCharFormat: TRtfCharFormat;
    FStartNumber: integer;
    procedure SetListType(const Value: TRtfListType);
  protected
    constructor Create(AParent: TRtfListItems; AType: TRtfListType; ADoc: TRtfDocument); reintroduce;

    function GetParagraph: TRtfParagraph; override;
    function GetNumber(ANumber: Integer): string;
    function GetAlphaNumber(ANumber: Integer; ALowercase: Boolean): string;
    function GetRomanNumber(ANumber: Integer; ALowercase: Boolean): string;

    function GetListCharFormat: TRtfCharFormat;
    function GetListCharacter(ANumber: Integer): string;

    function GetDefaultCharFormat: TRtfCharFormat; override;

    property NumberPrefix: string read FNumberPrefix write FNumberPrefix;
    property NumberSuffix: string read FNumberSuffix write FNumberSuffix;
    property NumberSeparator: string read FNumberSeparator write FNumberSeparator;
  public
    destructor Destroy; override;
    function AddList(AType: TRtfListType): TRtfListItems;
    function AddItem: TRtfListItem;

    function GetTotalTextIndent: Double;
    function Render: string; override;

    property ListIndent: Double read FListIndent write FListIndent;
    property Textindent: Double read FTextindent write FTextindent;
    property ListType: TRtfListType read FListType write SetListType;
    property StartNumber: integer read FStartNumber write FStartNumber;

    property DefaultCharFormat: TRtfCharFormat read GetDefaultCharFormat;
  end;

  TRtfList = class(TRtfListItems)
  public
    constructor Create(AType: TRtfListType; ADoc: TRtfDocument);

    function Render: string; override;
  end;


  TRtfParagraph = class(TRtfBlock)
  private
    FText: TStringBuilder;
    FLineSpacing: Double;
    FCharFormats: TRtfCharFormats;
    FAllowFootnote: Boolean;
    FAllowControlWord: Boolean;
    FFootNotes: TRtfFootNotes;
    FControlWords: TRtfFieldControlWords;
    FFirstLineindent: Double;
  protected
    type
      TToken = record
        Text: string;
        IsControl: Boolean;
      end;

      TTokenList = TSimpleLinkedList<TToken>;

      TDisjointRange = class
      public
        head: Integer;
        tail: Integer;
        Format: TRtfCharFormat;

        constructor Create(AHead, ATail: Integer; AFmt: TRtfCharFormat);
      end;

      TDisjointRanges = TObjectList<TDisjointRange>;

  protected
    function GetDefaultCharFormat: TRtfCharFormat; override;
    function BuildTokenList: TTokenList;
    function ExtractTokenList(AList: TTokenList): string;
  public
    constructor Create; reintroduce; overload;
    constructor Create(AAllowFootnote, AAllowControlWord: Boolean); reintroduce; overload;
    destructor Destroy; override;

    function Render: string; override;
    procedure SetText(AText: string);

    function AddCharFormat(AScopeBegin, AScopeEnd: Integer): TRtfCharFormat; overload;
    function AddCharFormat: TRtfCharFormat; overload;
    function AddFootnote(APosition: Integer): TRtfFootnote;
    procedure AddControlWord(APosition: Integer; AType: TRtfFieldControlWord.TFieldType);

    property Text: TStringBuilder read FText;
    property LineSpacing: Double read FLineSpacing write FLineSpacing;
    property FirstLineindent: Double read FFirstLineindent write FFirstLineindent;

  end;

  TRtfImage = class(TRtfBlock)
  private
    FImgType: TRtfImageFileType;
    FImgBytes: TBytes;
    FWidth: double;
    FHeight: double;
    FBmpWidth: integer;
    FBmpHeight: integer;
    FKeepAspectRatio: boolean;
    FStartNewPara: boolean;
    procedure SetHeight(const Value: double);
    procedure SetWidth(const Value: double);
  protected
    constructor Create(AFileName: string); reintroduce; overload;
    constructor Create(AImageStream: TStream); reintroduce; overload;
    constructor Create(AGraphic: TGraphic); reintroduce; overload;
    procedure Init;

    function FindGraphicClass(const ABuffer; const ABufferSize: Int64): TGraphicClass; overload;
    function FindGraphicClass(AExt: string): TGraphicClass; overload;
    function FindGraphicClass(AStream: TStream): TGraphicClass; overload;
    procedure BitmapToDIBStream(ABitmap: TBitmap; AStream: TStream);
    function ExtractImage: string;
  public
    destructor Destroy; override;
    function Render: string; override;

    property StartNewPara: boolean read FStartNewPara write FStartNewPara;
    property KeepAspectRatio: boolean read FKeepAspectRatio write FKeepAspectRatio;
    property Width: double read FWidth write SetWidth;
    property Height: double read FHeight write SetHeight;
  end;

  TRtfSection = class(TRtfBlock)
  private
    FSectionFooter: TRtfSectionFooter;
    FParentDocument: TRtfDocument;
    FPageOrientation: TRtfPaperOrientation;
    FStartEnd: TRtfSectionStartsEnd;
    FFooterPositionFromPageBottom: integer;
    FPageWidth: integer;
    FPageHeight: integer;
    function GetSectionFooter: TRtfSectionFooter;
  protected
    constructor Create(AType: TRtfSectionStartsEnd; ADoc: TRtfDocument); reintroduce;

    function GetDefaultCharFormat: TRtfCharFormat; override;

    property FooterPositionFromPageBottom: integer read FFooterPositionFromPageBottom write FFooterPositionFromPageBottom;
    property ParentDocument: TRtfDocument read FParentDocument write FParentDocument;
  public
    destructor Destroy; override;

    property StartEnd: TRtfSectionStartsEnd read FStartEnd;
    property PageOrientation: TRtfPaperOrientation read FPageOrientation write FPageOrientation;
    property SectionFooter: TRtfSectionFooter read GetSectionFooter;
    property PageWidth: integer read FPageWidth write FPageWidth;
    property PageHeight: integer read FPageHeight write FPageHeight;

    function Render: string; override;
  end;

  TRtfSectionFooter = class(TRtfBlocklist)
  protected
    constructor Create(AParent: TRtfSection);
  public
    function Render: string; override;
  end;

  TRtfTableCell = class(TRtfBlocklist)
  private
    FWidth: double;
    FHAlign: TRtfAlign;
    FVAlign: TRtfAlignVertical;
    FBorders: TRtfBorders;
    FMergeInfo: TRtfCellMergeInfo;
    FRowIndex: integer;
    FColIndex: integer;
    FParentTable: TRtfTable;
    FBackgroundColorIndex: TRtfColorDescriptor;
    FOuterLeftBorderClearance: double;
  protected
    constructor Create(AWidth: double; ARowIndex, AColIndex: integer; AParentTable: TRtfTable);

    function IsBeginOfColSpan: boolean;
    function IsBeginOfRowSpan: boolean;

    property MergeInfo: TRtfCellMergeInfo read FMergeInfo write FMergeInfo;
  public
    destructor Destroy; override;
    function IsMerged: boolean;
    procedure SetBorderColor(AColor: TRtfColorDescriptor);
    function Render: string; override;

    property Width: Double read FWidth write FWidth;
    property Borders: TRtfBorders read FBorders;
    property ParentTable: TRtfTable read FParentTable write FParentTable;
    property BackgroundColorIndex: TRtfColorDescriptor read FBackgroundColorIndex write FBackgroundColorIndex;
    property HAlign: TRtfAlign read FHAlign write FHAlign;
    property VAlign: TRtfAlignVertical read FVAlign write FVAlign;
    property RowIndex: integer read FRowIndex;
    property ColIndex: integer read FColIndex;
    property OuterLeftBorderClearance: double read FOuterLeftBorderClearance write FOuterLeftBorderClearance;
  end;

  TRtfTableCells = TObjectList<TRtfTableCell>;

  TRtfTable = class(TRtfBlock)
  private
    FRowCount: Integer;
    FColCount: Integer;
    FCells: array of array of TRtfTableCell;
    FDefaultCellWidth: Double;
    FRepresentativeList: TRtfTableCells;
    FRowHeight: array of Double;
    FRowKeepInSamePage: array of Boolean;
    FTitleRowCount: Integer;
    FFontSize: Double;
    FCellPadding: array of TRtfMargins;

    FHeaderBackgroundColour: TRtfColorDescriptor;
    FRowBackgroundColour: TRtfColorDescriptor;
    FRowAltBackgroundColour: TRtfColorDescriptor;
    function GetCellPadding(index: integer): TRtfMargins;
    function GetCell(ARow, ACol: integer): TRtfTableCell;
  protected
    function GetDefaultCharFormat: TRtfCharFormat; override;
    procedure ValidateAllMergedCellBorders;
    procedure ValidateMergedCellBorders(ARepresentative: TRtfTableCell);
    procedure ValidateMergedCellBorder(ARepresentative: TRtfTableCell; ADir: TRtfDirection);

  public
    constructor Create(ARowCount, AColCount: Integer; AHorizontalWidth, AFontSize: Double); reintroduce; virtual;
    destructor Destroy; override;

    function Render: string; override;

    procedure SetColWidth(ACol: integer; AWidth: Double);
    procedure SetRowHeight(ARow: integer; AHeight: Double);
    procedure SetRowKeepInSamePage(ARow: integer; AAllow: boolean);
    procedure SetInnerBorder(AStyle: TRtfBorderStyle; AWidth: double); overload;
    procedure SetInnerBorder(AStyle: TRtfBorderStyle; AWidth: double; AColor: TRtfColorDescriptor); overload;
    procedure SetOuterBorder(AStyle: TRtfBorderStyle; AWidth: double); overload;
    procedure SetOuterBorder(AStyle: TRtfBorderStyle; AWidth: double; AColor: TRtfColorDescriptor); overload;
    procedure SetHeaderBorderColors(AColorOuter, AColorInner: TRtfColorDescriptor);

    function Merge(ATopRow, ALeftCol, ARowSpan, AColSpan: integer): TRtfTableCell;

    property HeaderBackgroundColour: TRtfColorDescriptor read FHeaderBackgroundColour write FHeaderBackgroundColour;
    property RowBackgroundColour: TRtfColorDescriptor read FRowBackgroundColour write FRowBackgroundColour;
    property RowAltBackgroundColour: TRtfColorDescriptor read FRowAltBackgroundColour write FRowAltBackgroundColour;

    property RowCount: integer read FRowCount;
    property ColCount: integer read FColCount;
    property TitleRowCount: Integer read FTitleRowCount write FTitleRowCount;

    property CellPadding[index: integer]: TRtfMargins read GetCellPadding;
    property Cell[ARow, ACol: integer]: TRtfTableCell read GetCell;
  end;

  TRtfHeaderFooter = class(TRtfBlocklist)
  private
    //MagicWords ?
    FType: TRtfHeaderFooterType;
  public
    constructor Create(AType: TRtfHeaderFooterType);

    function Render: string; override;
  end;



  TRtfDocument = class(TRtfBlocklist)
  private
    FPaperSize: TRtfPaperSize;
    FPaperOrientation: TRtfPaperOrientation;
    FMargins: TRtfMargins;
    FLcid: TRtfLcid;
    FFontTable: TRtfFontTable;
    FColorTable: TRtfColorTable;
    FHeader: TRtfHeaderFooter;
    FFooter: TRtfHeaderFooter;
    procedure SetMargins(const Value: TRtfMargins);
    function GetFooter: TRtfHeaderFooter;
    function GetHeader: TRtfHeaderFooter;
  public
    constructor Create(APaperSize: TRtfPaperSize; APaperOrientation: TRtfPaperOrientation; ALcid: TRtfLcid);
    destructor Destroy; override;

    function Render: string; override;
    procedure ToStream(AStream: TStream);

    function DefaultColor: TRtfColorDescriptor;
    function DefaultFont: TRtfFontDescriptor;

    procedure SetDefaultFont(AFontName: string);
    function CreateFont(AFontName: string): TRtfFontDescriptor;

    function CreateColor(var AColor: TRtfColor): TRtfColorDescriptor; overload;
    function CreateColor(const AColor: TColor): TRtfColorDescriptor; overload;

    function AddTable(ARowCount, AColCount: Integer; AFontSize: Double): TRtfTable;

    property Margins: TRtfMargins read FMargins write SetMargins;
    property Header: TRtfHeaderFooter read GetHeader;
    property Footer: TRtfHeaderFooter read GetFooter;
  end;

  TRtfUtility = class
  public
    class function unicodeEncode(AString: string): string;
    class function pt2Twip(APt: Double): Integer;
    class function pt2HalfPt(APt: Double): Integer;

    class function paperDimensions(APaperSize: TRtfPaperSize): TSize;
    class function paperHeightInTwip(APaperSize: TRtfPaperSize; APaperOrientation: TRtfPaperOrientation): Integer;
    class function paperWidthInTwip(APaperSize: TRtfPaperSize; APaperOrientation: TRtfPaperOrientation): Integer;
    class function paperHeightInPt(APaperSize: TRtfPaperSize; APaperOrientation: TRtfPaperOrientation): double;
    class function paperWidthInPt(APaperSize: TRtfPaperSize; APaperOrientation: TRtfPaperOrientation): double;
  end;

implementation

uses
  Math, GIFImg, PNGImage, JPeg;

type
  TRtfDefaultValue = record
    FontSize: integer;
    Font: string;
    MarginLarge: double;
    MarginSmall: double;
  end;

var
  DefaultValue: TRtfDefaultValue;

{ TRtfDocument }

function TRtfDocument.AddTable(ARowCount, AColCount: Integer;
  AFontSize: Double): TRtfTable;
var
  LHorizontalWidth: Double;
begin
  LHorizontalWidth := TRtfUtility.paperWidthInPt(FPaperSize, FPaperOrientation)
     - FMargins[drLeft] - FMargins[drRight];
  Result := inherited AddTable(ARowCount, AColCount, LHorizontalWidth, AFontSize);
end;

constructor TRtfDocument.Create(APaperSize: TRtfPaperSize;
  APaperOrientation: TRtfPaperOrientation; ALcid: TRtfLcid);
begin
  inherited Create;
  FPaperSize := APaperSize;
  FPaperOrientation := APaperOrientation;

  FMargins := TRtfMargins.Create;
  if APaperOrientation = poPortrait then
  begin
    FMargins[drTop] := DefaultValue.MarginSmall;
    FMargins[drRight] := DefaultValue.MarginLarge;
    FMargins[drBottom] := DefaultValue.MarginSmall;
    FMargins[drLeft] := DefaultValue.MarginLarge;
  end
  else
  begin
    FMargins[drTop] := DefaultValue.MarginLarge;
    FMargins[drRight] := DefaultValue.MarginSmall;
    FMargins[drBottom] := DefaultValue.MarginLarge;
    FMargins[drLeft] := DefaultValue.MarginSmall;
  end;

  FLcid := ALcid;
  FFontTable := TRtfFontTable.Create;
  FFontTable.Add(DefaultValue.Font);
  FColorTable := TRtfColorTable.Create;
  FColorTable.Add(TRtfColor.Create);

  FHeader := nil;
  FFooter := nil;
end;

function TRtfDocument.CreateColor(var AColor: TRtfColor): TRtfColorDescriptor;
begin
  Result := FColorTable.IndexOfColor(AColor);
end;

function TRtfDocument.CreateColor(const AColor: TColor): TRtfColorDescriptor;
var
  LRTFColor: TRtfColor;
begin
  LRTFColor := TRtfColor.Create(AColor);
  Result := FColorTable.IndexOfColor(LRTFColor);
end;

function TRtfDocument.CreateFont(AFontName: string): TRtfFontDescriptor;
begin
  Result := FFontTable.IndexOf(AFontName);
  if Result = -1 then
  begin
    Result := FFontTable.Add(AFontName);
  end;
end;

function TRtfDocument.DefaultColor: TRtfColorDescriptor;
begin
  Result := 0;
end;

function TRtfDocument.DefaultFont: TRtfFontDescriptor;
begin
  Result := 0;
end;

destructor TRtfDocument.Destroy;
begin
  FreeAndNil(FFontTable);
  FreeAndNil(FColorTable);
  inherited;
end;

function TRtfDocument.GetFooter: TRtfHeaderFooter;
begin
  if not Assigned(FFooter) then
  begin
    FFooter := TRtfHeaderFooter.Create(hfFooter);
  end;
  Result := FFooter;
end;

function TRtfDocument.GetHeader: TRtfHeaderFooter;
begin
  if not Assigned(FHeader) then
  begin
    FHeader := TRtfHeaderFooter.Create(hfHeader);
  end;
  Result := FHeader;
end;

function TRtfDocument.Render: string;
var
  LStringBuilder: TStringBuilder;
  i: Integer;
  LColor: TRtfColor;
begin
  LStringBuilder := TStringBuilder.Create('{\rtf1\ansi\deff0');
  try
    LStringBuilder.AppendLine;

    LStringBuilder.AppendLine('{\fonttbl');
    for i := 0 to FFontTable.Count -1 do
    begin
      LStringBuilder.AppendFormat('{\f%d %s;}', [i, TRtfUtility.unicodeEncode(FFontTable[i]) ]);
      LStringBuilder.AppendLine;
    end;
    LStringBuilder.AppendLine('}');
    LStringBuilder.AppendLine;

    LStringBuilder.AppendLine('{\colortbl');
    LStringBuilder.AppendLine(';');   // default color is only a ;
    for i := 1 to FColorTable.Count -1 do
    begin
      LColor := FColorTable[i];
      LStringBuilder.AppendFormat('\red%s\green%s\blue%s;', [LColor.Red, LColor.Green, LColor.Blue]);
      LStringBuilder.AppendLine;
    end;
    LStringBuilder.AppendLine('}');
    LStringBuilder.AppendLine;

    // Preliminary
    LStringBuilder.AppendLine('\deflang' + IntToStr(Integer(FLcid)) + '\plain\fs'
      + IntToStr(TRtfUtility.pt2HalfPt(DefaultValue.FontSize)) + '\widowctrl\hyphauto\ftnbj');
    // page size
    LStringBuilder.AppendLine('\paperw' + IntToStr(TRtfUtility.paperWidthInTwip(FPaperSize, FPaperOrientation))
      + '\paperh' + IntToStr(TRtfUtility.paperHeightInTwip(FPaperSize, FPaperOrientation)));
    // page margin
    LStringBuilder.AppendLine('\margt' + IntToStr(TRtfUtility.pt2Twip(FMargins[drTop])));
    LStringBuilder.AppendLine('\margr' + IntToStr(TRtfUtility.pt2Twip(FMargins[drRight])));
    LStringBuilder.AppendLine('\margb' + IntToStr(TRtfUtility.pt2Twip(FMargins[drBottom])));
    LStringBuilder.AppendLine('\margl' + IntToStr(TRtfUtility.pt2Twip(FMargins[drLeft])));
    // orientation
    if (FPaperOrientation = poLandscape) then
    begin
        LStringBuilder.AppendLine('\landscape');
    end;
    // header/footer
    if Assigned(FHeader) then
    begin
      LStringBuilder.Append(FHeader.Render());
    end;
    if Assigned(FFooter) then
    begin
      LStringBuilder.Append(FFooter.Render());
    end;
    LStringBuilder.AppendLine();

    LStringBuilder.AppendLine( inherited Render );
    LStringBuilder.AppendLine('}');
    Result := LStringBuilder.ToString;
  finally
    LStringBuilder.Free;
  end;
end;

procedure TRtfDocument.SetDefaultFont(AFontName: string);
begin
  FFontTable[0] := AFontName;
end;

procedure TRtfDocument.SetMargins(const Value: TRtfMargins);
begin
  FMargins.CopyFrom(Value);
end;

procedure TRtfDocument.ToStream(AStream: TStream);
var
  LTextStream: TStringStream;
begin
  LTextStream := TStringStream.Create(Render);
  try
    AStream.CopyFrom(LTextStream, 0);
  finally
    LTextStream.Free;
  end;
end;

{ TRtfBlocklist }

constructor TRtfBlocklist.Create;
begin
  Create(True, True, True, True, True, True, True);
end;

constructor TRtfBlocklist.Create(AAllowParagraph, AAllowTable: Boolean);
begin
  Create(AAllowParagraph, True, True, True, AAllowTable, True, True);
end;

procedure TRtfBlocklist.AddBlock(ABlock: TRtfBlock);
begin
  if Assigned(ABlock) then
    FBlocks.Add(ABlock);
end;

function TRtfBlocklist.AddImage(AImageFilename: string): TRtfImage;
begin
  if ( not FAllowImage) then begin
    raise TRtfException.Create('Image is not allowed.');
  end;
  Result := TRtfImage.Create(AImageFilename);
  addBlock(Result);
end;

function TRtfBlocklist.AddImage(AImageStream: TStream): TRtfImage;
begin
  if ( not FAllowImage) then begin
    raise TRtfException.Create('Image is not allowed.');
  end;
  Result := TRtfImage.Create(AImageStream);
  addBlock(Result);
end;

function TRtfBlocklist.AddImage(AImageGraphic: TGraphic): TRtfImage;
begin
  if ( not FAllowImage) then begin
    raise TRtfException.Create('Image is not allowed.');
  end;
  Result := TRtfImage.Create(AImageGraphic);
  addBlock(Result);
end;

function TRtfBlocklist.AddList(AType: TRtfListType;
  ADoc: TRtfDocument): TRtfList;
begin
  if not FAllowList then
    raise TRtfException.Create('List is not allowed.');
  Result := TRtfList.Create(AType, ADoc);
  AddBlock(Result);
end;

function TRtfBlocklist.AddParagraph: TRtfParagraph;
begin
  if not FAllowParagraph then
    raise TRtfException.Create('Paragraph is not allowed.');
  Result := TRtfParagraph.Create(FAllowFootNote, FAllowControlWord);
  AddBlock(Result);
end;

function TRtfBlocklist.AddSection(AType: TRtfSectionStartsEnd;
  ADoc: TRtfDocument): TRtfSection;
begin
  Result := TRtfSection.Create(AType, ADoc);
  AddBlock(Result);
end;

procedure TRtfBlocklist.AddSpecialInsert(AInsert: TRtfSpecialInsert);
begin
  if not FAllowSpecialInsert then
    raise TRtfException.Create('Special Insert is not allowed.');
  if Assigned(AInsert) then
    AddBlock(AInsert);
end;

function TRtfBlocklist.AddTable(ARowCount, AColCount: Integer; AHorizontalWidth,
  AFontSize: Double): TRtfTable;
begin
  if not FAllowTable then
    raise TRtfException.Create('Table is not allowed.');
  Result := TRtfTable.Create(ARowCount, AColCount, AHorizontalWidth, AFontSize);
  AddBlock(Result);
end;

constructor TRtfBlocklist.Create(AAllowParagraph, AAllowFootnote,
  AAllowControlWord, AAllowImage, AAllowTable, AAllowSpecialInsert, AAllowList: Boolean);
begin
  inherited Create;
  FBlocks := TObjectList<TRtfBlock>.Create;
  FAllowParagraph := AAllowParagraph;
  FAllowFootnote := AAllowFootnote;
  FAllowControlWord := AAllowControlWord;
  FAllowImage := AAllowImage;
  FAllowTable := AAllowTable;
  FAllowSpecialInsert := AAllowSpecialInsert;
  FAllowList := AAllowList;
  FDefaultCharFormat := nil;
end;

destructor TRtfBlocklist.Destroy;
begin
  FreeAndNil(FBlocks);
  inherited;
end;

function TRtfBlocklist.GetDefaultCharFormat: TRtfCharFormat;
begin
  if not Assigned(FDefaultCharFormat) then
  begin
    FDefaultCharFormat := TRtfCharFormat.Create(-1, -1, 0);
  end;
  Result := FDefaultCharFormat;
end;

function TRtfBlocklist.Render: string;
var
  LStringBuilder: TStringBuilder;
  i: Integer;
begin
  LStringBuilder := TStringBuilder.Create;
  try
    LStringBuilder.AppendLine;
    for i := 0 to FBlocks.Count -1 do
    begin
      if Assigned(FDefaultCharFormat) and Assigned(FBlocks[i].DefaultCharFormat) then
        FBlocks[i].DefaultCharFormat.CopyFrom(FDefaultCharFormat);
      LStringBuilder.AppendLine(FBlocks[i].Render);
    end;

    Result := LStringBuilder.ToString;
  finally
    LStringBuilder.Free;
  end;
end;

procedure TRtfBlocklist.TransferBlocksTo(ATarget: TRtfBlocklist);
var
  i: Integer;
begin
  for i := 0 to FBlocks.Count -1 do
  begin
    ATarget.AddBlock(FBlocks[i]);
  end;
  FBlocks.OwnsObjects := False;
  try
    FBlocks.Clear;
  finally
    FBlocks.OwnsObjects := True;
  end;
end;

{ TRtfFontStyle }

constructor TRtfFontStyle.Create;
begin
  inherited;
  FStyleAdd := [];
  FStyleRemove := [];
end;

procedure TRtfFontStyle.AddStyle(AStyle: TRtfFontStyleFlag);
begin
  include(FStyleAdd, AStyle);
  exclude(FStyleRemove, AStyle);
end;

procedure TRtfFontStyle.CopyFrom(ASource: TRtfFontStyle);
begin
  if Assigned(ASource) then
  begin
    FStyleAdd := ASource.StyleAdd;
    FStyleRemove := ASource.StyleRemove;
  end;
end;

constructor TRtfFontStyle.Create(AStyle: TRtfFontStyle);
begin
  Create;
  FStyleAdd := AStyle.StyleAdd;
  FStyleRemove := AStyle.StyleRemove;
end;

function TRtfFontStyle.IsEmpty: boolean;
begin
  Result := (FStyleAdd = []) and (FStyleRemove = []);
end;

procedure TRtfFontStyle.RemoveStyle(AStyle: TRtfFontStyleFlag);
begin
  exclude(FStyleAdd, AStyle);
  include(FStyleRemove, AStyle);
end;

{ TRtfMargins }

constructor TRtfMargins.Create;
begin
  Create(0, 0, 0, 0);
end;

procedure TRtfMargins.CopyFrom(ASource: TRtfMargins);
begin
  if Assigned(ASource) then
  begin
    FMargins[drTop] := ASource.Margins[drTop];
    FMargins[drRight] := ASource.Margins[drRight];
    FMargins[drBottom] := ASource.Margins[drBottom];
    FMargins[drLeft] := ASource.Margins[drLeft];
  end;
end;

constructor TRtfMargins.Create(ATop, ARight, ABottom, ALeft: double);
begin
  inherited Create;
  FMargins[drTop] := ATop;
  FMargins[drRight] := ARight;
  FMargins[drBottom] := ABottom;
  FMargins[drLeft] := ALeft;
end;

function TRtfMargins.Equals(Obj: TObject): boolean;
begin
  if Obj is TRtfMargins then
  begin
    Result :=
      SameValue(FMargins[drTop], TRtfMargins(Obj)[drTop]) and
      SameValue(FMargins[drRight], TRtfMargins(Obj)[drRight]) and
      SameValue(FMargins[drBottom], TRtfMargins(Obj)[drBottom]) and
      SameValue(FMargins[drLeft], TRtfMargins(Obj)[drLeft]);
  end
  else
    Result := inherited;
end;

function TRtfMargins.GetMargins(index: TRtfDirection): double;
begin
  Result := FMargins[index];
end;

procedure TRtfMargins.SetMargins(index: TRtfDirection; const Value: double);
begin
  FMargins[index] := Value;
end;

{ TRtfBorder }

constructor TRtfBorder.Create;
begin
  inherited Create;
  FStyle := bsNone;
  FWidth := 0.5;
  FColorDesc := 0;
end;

function TRtfBorder.Equals(Obj: TObject): boolean;
begin
  if Obj is TObject then
  begin
    Result :=
      (FStyle = TRtfBorder(Obj).Style) and
      (FColorDesc = TRtfBorder(Obj).ColorDesc) and
      SameValue(FWidth, TRtfBorder(Obj).Width);
  end
  else
    Result := inherited;
end;

function TRtfBorder.GetHashCode: integer;
begin
  Result := trunc(FWidth * 1000) * 10 + ord(FStyle);
end;

{ TRtfBorders }

constructor TRtfBorders.Create;
begin
  inherited Create;
  FBorders[drTop] := TRtfBorder.Create;
  FBorders[drRight] := TRtfBorder.Create;
  FBorders[drBottom] := TRtfBorder.Create;
  FBorders[drLeft] := TRtfBorder.Create;
end;

destructor TRtfBorders.Destroy;
begin
  FBorders[drTop].Free;
  FBorders[drRight].Free;
  FBorders[drBottom].Free;
  FBorders[drLeft].Free;
  inherited;
end;

function TRtfBorders.GetBorders(index: TRtfDirection): TRtfBorder;
begin
  Result := FBorders[index];
end;

{ TRtfColor }

constructor TRtfColor.Create(ARed, AGreen, ABlue: byte);
begin
  inherited Create;
  FColor := ARed shl 16 + AGreen shl 8 + ABlue;
end;

constructor TRtfColor.Create;
begin
  Create(0, 0, 0);
end;

constructor TRtfColor.Create(AColor: TColor);
begin
  AColor := ColorToRGB(AColor);

  Create(GetRValue(AColor), GetGValue(AColor), GetBValue(AColor));
end;

constructor TRtfColor.Create(AHex: string);
begin
  inherited Create;
  if (Length(AHex) = 6) then
  begin
    FColor := StrToIntDef('$' + AHex, -1);
    if FColor = -1 then
    begin
      FColor := 0;
    end;
  end;
end;

function TRtfColor.Equals(Obj: TObject): boolean;
begin
  if Obj is TRtfColor then
  begin
    Result := FColor = TRtfColor(Obj).Color;
  end
  else
    Result := inherited;
end;

function TRtfColor.GetBlue: string;
begin
  Result := IntToStr(FColor and $ff);
end;

function TRtfColor.GetGreen: string;
begin
  Result := IntToStr((FColor and $ff00) shr 8);
end;

function TRtfColor.GetHashCode: integer;
begin
  Result := FColor;
end;

function TRtfColor.GetRed: string;
begin
  Result := IntToStr((FColor and $ff0000) shr 16);
end;

{ TRtfCellMergeInfo }

constructor TRtfCellMergeInfo.Create(ARepresentative: TRtfTableCell; ARowSpan,
  AColSpan, ARowIndex, AColIndex: integer);
begin
  inherited Create;
  FRepresentative := ARepresentative;
  FRowSpan := ARowSpan;
  FColSpan := AColSpan;
  FRowIndex := ARowIndex;
  FColIndex := AColIndex;
end;


{ TRtfBlock }

function TRtfBlock.AlignmentCode: string;
begin
  case FAlignment of
    alLeft: Result := '\ql';
    alRight: Result := '\qr';
    alCenter: Result := '\qc';
    alFullyJustified: Result := '\qj';
  else
    Result := '\qd';
  end;
end;

constructor TRtfBlock.Create;
begin
  inherited Create;
  FAlignment := alNone;
  FBlockTail := '';
  FStartNewPage := False;
  FBlockHead := '';
  FMargins := TRtfMargins.Create;
  FDefaultCharFormat := nil;
end;

destructor TRtfBlock.Destroy;
begin
  FreeAndNil(FMargins);
  FreeAndNil(FDefaultCharFormat);
  inherited;
end;

function TRtfBlock.GetDefaultCharFormat: TRtfCharFormat;
begin
  Result := nil;
end;

{ TRtfCharFormat }

procedure TRtfCharFormat.CopyFrom(ASource: TRtfCharFormat);
begin
  if Assigned(ASource) then
  begin
    FScopeBegin := ASource.ScopeBegin;
    FScopeEnd := ASource.ScopeEnd;
    if (ASource.Font <> -1) and (Font = -1) then FFont := ASource.Font;
    if (ASource.AnsiFont <> -1) and (AnsiFont = -1) then FAnsiFont := ASource.AnsiFont;
    if (ASource.FontSize <> -1) and (FontSize = -1) then FFontSize := ASource.FontSize;
    if not ASource.FontStyle.IsEmpty and FFontStyle.IsEmpty then FFontStyle.CopyFrom(ASource.FontStyle);

    if (ASource.BGColor <> -1) and (BGColor = -1) then FBGColor := ASource.BGColor;
    if (ASource.FGColor <> -1) and (FGColor = -1) then FFGColor := ASource.FGColor;
  end;
end;

constructor TRtfCharFormat.Create(AScopeBegin, AScopeEnd, ATextLength: integer);
begin
  inherited Create;
  FScopeBegin := -1;
  FScopeEnd := -1;
  FFont := -1;
  FAnsiFont := -1;
  FFontSize := -1;
  FFontStyle := TRtfFontStyle.Create;
  FBGColor := -1;
  FFGColor := -1;
  FTwoInOneStyle := toNotEnabled;
  FBookmark := '';
  SetRange(AScopeBegin, AScopeEnd, ATextLength);
end;

destructor TRtfCharFormat.Destroy;
begin
  FreeAndNil(FFontStyle);
  inherited;
end;

function TRtfCharFormat.RenderHead: string;
const
  CFONTSTYLEMAP: array[TRtfFontStyleFlag] of string = ('b', 'i', 'ul', 'super', 'sub', 'scaps', 'strike');
var
  LStringBuilder: TStringBuilder;
  LFontStyle: TRtfFontStyleFlag;
begin
  LStringBuilder := TStringBuilder.Create('{');
  try
    if FLocalHyperlink <> '' then
    begin
      LStringBuilder.Append('{\field{\*\fldinst HYPERLINK \\l ');
      LStringBuilder.Append('"' + FLocalHyperlink + '"');
      if FLocalHyperlinkTip <> '' then
      begin
        LStringBuilder.Append(' \\o "' + FLocalHyperlinkTip + '"');
      end;
      LStringBuilder.Append('}{\fldrslt{');
    end;
    if (FFont <> -1) or (FAnsiFont <> -1) then
    begin
      if (FFont = -1) then
      begin
        LStringBuilder.AppendFormat('\f%d', [FAnsiFont]);
      end
      else if (FAnsiFont = -1) then
      begin
        LStringBuilder.AppendFormat('\f%d', [FFont]);
      end
      else
      begin
        LStringBuilder.AppendFormat('\loch\af%d\hich\af%d\dbch\af%d', [FAnsiFont, FAnsiFont, FFont]);
      end;
    end;
    if FFontSize > 0 then
    begin
      LStringBuilder.AppendFormat('\fs%d', [TRtfUtility.pt2HalfPt(FFontSize)]);
    end;
    if FFGColor > 0 then
    begin
      LStringBuilder.AppendFormat('\cf%d', [FFGColor]);
    end;
    if FBGColor > 0 then
    begin
      LStringBuilder.AppendFormat('\chshdng0\chcbpat%d\cb%d', [FBGColor, FBGColor]);
    end;

    for LFontStyle := Low(TRtfFontStyleFlag) to High(TRtfFontStyleFlag) do
    begin
      if LFontStyle in FFontStyle.StyleAdd then
      begin
        LStringBuilder.Append('\' + CFONTSTYLEMAP[LFontStyle]);
      end
      else if LFontStyle in FFontStyle.StyleRemove then
      begin
        LStringBuilder.Append('\' + CFONTSTYLEMAP[LFontStyle] + '0');
      end;
    end;
    if FTwoInOneStyle <> toNotEnabled then
    begin
      LStringBuilder.Append('\twoinone');
      case FTwoInOneStyle of
        toNone: LStringBuilder.Append('0');
        toParentheses: LStringBuilder.Append('1');
        toSquareBrackets: LStringBuilder.Append('2');
        toAngledBrackets: LStringBuilder.Append('3');
        toBraces: LStringBuilder.Append('4');
      end;
    end;
    Result := LStringBuilder.ToString;
    if pos('\', Result) > 0 then
    begin
      LStringBuilder.Append(' ');
    end;
    if FBookmark <> '' then
    begin
      LStringBuilder.Append('{\*\bkmkstart ' + FBookmark + '}');
    end;
    Result := LStringBuilder.ToString;
  finally
    LStringBuilder.Free;
  end;
end;

function TRtfCharFormat.RenderTail: string;
var
  LStringBuilder: TStringBuilder;
begin
  LStringBuilder := TStringBuilder.Create;
  try
    if FBookmark <> '' then
    begin
      LStringBuilder.Append('{\*\bkmkend ' + FBookmark + '}');
    end;
    if FLocalHyperlink <> '' then
    begin
      LStringBuilder.Append('}}}');
    end;
    LStringBuilder.Append('}');
    Result := LStringBuilder.ToString;
  finally
    LStringBuilder.Free;
  end;
end;

procedure TRtfCharFormat.SetFont(ARtfDocument: TRtfDocument; AFont: TFont);
var
  LFontNameDescriptor: TRtfFontDescriptor;
  LFontColorDescriptor: TRtfColorDescriptor;
begin
  LFontNameDescriptor := ARtfDocument.CreateFont(AFont.Name);
  LFontColorDescriptor := ARtfDocument.CreateColor(AFont.Color);
  FFont := LFontNameDescriptor;
  FFGColor := LFontColorDescriptor;
  FFontSize := AFont.Size;

  // TFontStyle = (fsBold, fsItalic, fsUnderline, fsStrikeOut);
  if TFontStyle.fsBold in AFont.Style then
    FFontStyle.AddStyle( TRtfFontStyleFlag.fsBold);
  if TFontStyle.fsItalic in AFont.Style then
    FFontStyle.AddStyle( TRtfFontStyleFlag.fsItalic);
  if TFontStyle.fsUnderline in AFont.Style then
    FFontStyle.AddStyle( TRtfFontStyleFlag.fsUnderline);
  if TFontStyle.fsStrikeOut in AFont.Style then
    FFontStyle.AddStyle( TRtfFontStyleFlag.fsStrike);
end;

procedure TRtfCharFormat.SetRange(AScopeBegin, AScopeEnd, ATextLength: integer);
begin
  if AScopeBegin > AScopeEnd then
  begin
    raise TRtfException.CreateFmt('Invalid range: (%d,%d)', [AScopeBegin, AScopeEnd]);
  end
  else if (AScopeBegin < 0) or (AScopeEnd < 0) then
  begin
    if (AScopeBegin <> -1) or (AScopeEnd <> -1) then
    begin
      raise TRtfException.CreateFmt('Invalid range: (%d,%d)', [AScopeBegin, AScopeEnd]);
    end;
  end;
  if (AScopeEnd >= ATextLength) then
  begin
    raise TRtfException.CreateFmt('Range ending out of range: %d', [AScopeEnd]);
  end;
  FScopeBegin := AScopeBegin;
  FScopeEnd := AScopeEnd;
end;

{ TRtfUtility }

class function TRtfUtility.paperDimensions(APaperSize: TRtfPaperSize): TSize;
begin
  case APaperSize of
    psA5:
      begin
        Result.cx := 8391;
        Result.cy := 11906;
      end;
    psA4:
      begin
        Result.cx := 11906;
        Result.cy := 16838;
      end;
    psA3:
      begin
        Result.cx := 16838;
        Result.cy := 23811;
      end;
    psLetter:
      begin
        Result.cx := 15840;
        Result.cy := 12240;
      end;
    psLegal:
      begin
        Result.cx := 20212;
        Result.cy := 12240;
      end;
  end;
end;

class function TRtfUtility.paperHeightInPt(APaperSize: TRtfPaperSize;
  APaperOrientation: TRtfPaperOrientation): double;
begin
  Result := paperHeightInTwip(APaperSize, APaperOrientation) / 20;
end;

class function TRtfUtility.paperHeightInTwip(APaperSize: TRtfPaperSize;
  APaperOrientation: TRtfPaperOrientation): Integer;
var
  LSize: TSize;
begin
  Result := -1;
  LSize := paperDimensions(APaperSize);
  case APaperOrientation of
    poPortrait:
    begin
      if LSize.cx < LSize.cy then
        Result := LSize.cy
      else
        Result := LSize.cx;
    end;
    poLandscape:
    begin
      if LSize.cx < LSize.cy then
        Result := LSize.cx
      else
        Result := LSize.cy;
    end;
  end;
end;

class function TRtfUtility.paperWidthInPt(APaperSize: TRtfPaperSize;
  APaperOrientation: TRtfPaperOrientation): double;
begin
  Result := paperWidthInTwip(APaperSize, APaperOrientation) / 20;
end;

class function TRtfUtility.paperWidthInTwip(APaperSize: TRtfPaperSize;
  APaperOrientation: TRtfPaperOrientation): Integer;
var
  LSize: TSize;
begin
  Result := -1;
  LSize := paperDimensions(APaperSize);
  case APaperOrientation of
    poPortrait:
    begin
      if LSize.cx < LSize.cy then
        Result := LSize.cx
      else
        Result := LSize.cy;
    end;
    poLandscape:
    begin
      if LSize.cx < LSize.cy then
        Result := LSize.cy
      else
        Result := LSize.cx;
    end;
  end;

end;

class function TRtfUtility.pt2HalfPt(APt: Double): integer;
begin
  Result := Trunc(APt * 2);
end;

class function TRtfUtility.pt2Twip(APt: Double): Integer;
begin
  if IsNan(APt) then
    Result := 0
  else
    Result := Trunc(APt * 20);
end;

class function TRtfUtility.unicodeEncode(AString: string): string;
var
  LStringBuilder: TStringBuilder;
  i: Integer;
  LUniCode: Integer;
begin
  LStringBuilder := TStringBuilder.Create;
  try
    for i := 1 to Length(AString) do
    begin
      LUniCode := Integer(AString[i]);
      if AString[i] = #10 then
      begin
        LStringBuilder.AppendLine('\line');
      end
      else if AString[i] = #13 then
      begin

      end
      else if AString[i] = #9 then
      begin
        LStringBuilder.Append('\tab');
      end
      else if LUniCode <= $ff then
      begin
        if (LUniCode = $5c) or (LUniCode = $7b) or (LUniCode = $7d) then
        begin
          LStringBuilder.AppendFormat('\''%2x', [LUniCode]);
        end
        else if (LUniCode >= 0) and (LUniCode < $20) then
        begin
          LStringBuilder.AppendFormat('\''%2x', [LUniCode]);
        end
        else if (LUniCode >= 20) and (LUniCode < $80) then
        begin
          LStringBuilder.Append(AString[i]);
        end
        else
        begin
          LStringBuilder.AppendFormat('\''%2x', [LUniCode]);
        end;
      end
      else if LUniCode <= $8000 then
      begin
        LStringBuilder.AppendFormat('\uc1\u%d*', [LUniCode]);
      end
      else if LUniCode <= $ffff then
      begin
        LStringBuilder.AppendFormat('\uc1\u%d*', [LUniCode - $10000]);
      end
      else
      begin
        LStringBuilder.Append('\uc1\u9633*');
      end;
    end;
    Result := LStringBuilder.ToString;
  finally
    LStringBuilder.Free;
  end;
end;

{ TRtfHeaderFooter }

constructor TRtfHeaderFooter.Create(AType: TRtfHeaderFooterType);
begin
  inherited Create(True, False, True, True, False, True, False);
  FType := AType;
end;

function TRtfHeaderFooter.Render: string;
var
  LStringBuilder: TStringBuilder;
  i: Integer;
begin
  LStringBuilder := TStringBuilder.Create;
  try
    case FType of
      hfHeader: LStringBuilder.AppendLine('{\header');
      hfFooter: LStringBuilder.AppendLine('{\footer');
    end;
    LStringBuilder.AppendLine();
    for i := 0 to FBlocks.Count -1 do
    begin
      if Assigned(FDefaultCharFormat) and (Assigned(FBlocks[i].FDefaultCharFormat)) then
      begin
        FBlocks[i].FDefaultCharFormat.CopyFrom(FDefaultCharFormat);
      end;
      LStringBuilder.AppendLine(FBlocks[i].Render);
    end;
    LStringBuilder.AppendLine('}');
    Result := LStringBuilder.ToString;
  finally
    LStringBuilder.Free;
  end;
end;

{ TRtfColorTable }

function TRtfColorTable.IndexOfColor(var AColor: TRtfColor): TRtfColorDescriptor;
var
  i: Integer;
begin
  Result := IndexOf(AColor);
  if Result = -1 then
  begin
    // entry 0 is default color, ignore in search
    for i := 1 to Count -1 do
    begin
      if Items[i].Equals(AColor) then
      begin
        Result := i;
        AColor.Free;
        AColor := Items[i];
        Exit;
      end;
    end;
    Result := Add(AColor);
  end;
end;

{ TRtfParagraph }

constructor TRtfParagraph.Create;
begin
  Create(False, False);
end;

function TRtfParagraph.AddCharFormat(AScopeBegin,
  AScopeEnd: Integer): TRtfCharFormat;
begin
  Result := TRtfCharFormat.Create(AScopeBegin, AScopeEnd, FText.Length);
  FCharFormats.Add(Result);
end;

function TRtfParagraph.AddCharFormat: TRtfCharFormat;
begin
  Result := AddCharFormat(-1, -1);
end;

procedure TRtfParagraph.AddControlWord(APosition: Integer;
  AType: TRtfFieldControlWord.TFieldType);
var
  LControlWord: TRtfFieldControlWord;
  i: integer;
begin
  if not FAllowControlWord then
    raise TRtfException.Create('ControlWord is not allowed.');
  LControlWord := TRtfFieldControlWord.Create(APosition, AType);
  for i := 0 to FControlWords.Count -1 do
  begin
    if FControlWords[i].Position = APosition then
    begin
      FControlWords[i] := LControlWord;
      Exit;
    end;
  end;
  FControlWords.Add(LControlWord);
end;

function TRtfParagraph.AddFootnote(APosition: Integer): TRtfFootnote;
begin
  if not FAllowFootnote then
    raise TRtfException.Create('Footnote is not allowed.');
  Result := TRtfFootNote.Create(APosition, FText.Length);
  FFootNotes.Add(Result);
end;

function TRtfParagraph.BuildTokenList: TTokenList;
var
  LNode, newNode: TTokenList.TNode;
  LToken, newTok, tok, newTok1, newTok2, newTok3: TToken;
  LDRanges: TDisjointRanges;
  i,j, index, count: Integer;

  fmt: TRtfCharFormat;
  range: TDisjointRange;

  delList: TDisjointRanges;
  addList: TDisjointRanges;
  addAnchorList: TDisjointRanges;
  LRange, LNewRange: TDisjointRange;
  LPos: integer;
begin
  Result := TTokenList.Create;

  // Transform possibly overlapped character format ranges into
  // disjoint ranges.
  LDRanges := TDisjointRanges.Create;
  try
    for i := 0 to FCharFormats.Count -1 do
    begin
      fmt := FCharFormats[i];
      range := nil;
      if (fmt.ScopeBegin = -1) and (fmt.ScopeEnd = -1) then
      begin
        range := TDisjointRange.Create(1, FText.Length, fmt);
      end
      else if (fmt.ScopeBegin <= fmt.ScopeEnd) then
      begin
        range := TDisjointRange.Create(fmt.ScopeBegin, fmt.ScopeEnd, fmt);
      end
      else
        continue;

      // make the ranges disjoint from each other.
      delList := TDisjointRanges.Create(False);
      addList := TDisjointRanges.Create(False);
      addAnchorList := TDisjointRanges.Create(False);
      for j := 0 to LDRanges.Count -1 do
      begin
        LRange := LDRanges[j];
        if ((range.head <= LRange.head) and (range.tail >= LRange.tail)) then
        begin
          // former range is totally covered by the later
          //       |--------| r
          //   |-----------------| range
          delList.Add(LRange);
        end
        else if ((range.head <= LRange.head) and (range.tail >= LRange.head) and (range.tail < LRange.tail)) then
        begin
          // former range is partially covered
          //          |------------------| r
          //     |-----------------| range
          LRange.head := range.tail + 1;
        end else if ((range.head > LRange.head) and (range.head <= LRange.tail) and (range.tail >= LRange.tail)) then
        begin
          // former range is partially covered
          //     |------------------| r
          //          |-----------------| range
          LRange.tail := range.head - 1;
        end else if ((range.head > LRange.head) and (range.tail < LRange.tail)) then
        begin
          // later range is totally covered by the former
          //   |----------------------| r
          //        |---------| range
          LNewRange := TDisjointRange.Create(range.tail + 1, LRange.tail, LRange.Format);
          LRange.tail := range.head - 1;
          addList.Add(LNewRange);
          addAnchorList.Add(LRange);
        end;
      end;
      LDRanges.Add(range);
      for j := 0 to delList.Count-1 do
      begin
        LDRanges.Extract(delList[j]);
      end;
      for j := 0 to addList.Count-1 do begin
        index := LDRanges.IndexOf(addAnchorList[j]);
        if (index < 0) then begin
            continue;
        end;
        LDRanges.Insert(index, addList[j]);
      end;
    end;

    LToken.Text := FText.ToString;
    LToken.IsControl := False;
    Result.AddLast(LToken);

    // --------------------------------------------------
    // Build token list from head[] and tail[].
    // --------------------------------------------------
    for i := 0 to LDRanges.Count-1 do begin
      LRange := LDRanges[i];
      count := 0;
      // process head[i]
      if (LRange.head = 1) then begin
        newTok.isControl := true;
        newTok.text := LRange.format.renderHead();
        Result.AddFirst(newTok);
      end else begin
        LNode := Result.First;
        while Assigned(LNode) do
        begin
          tok := LNode.Value;

          if (not tok.isControl) then begin
            Count := Count + Length(tok.Text);
            if ((Count -1) = LRange.head) then begin
              newTok.isControl := true;
              newTok.text := LRange.format.renderHead();
              while Assigned(LNode.Next) and (LNode.Next.Value.isControl) do
              begin
                LNode := LNode.Next;
              end;
              Result.AddAfter(LNode, newTok);
              break;
            end else if ((count -1) > LRange.head) then
            begin
              newTok1.isControl := false;
              newTok1.text := Copy(tok.Text, 1, Length(tok.Text) - (count - LRange.head) + 1 -2);  // +1?
              newNode := Result.AddAfter(LNode, newTok1);
              newTok2.isControl := true;
              newTok2.text := LRange.format.renderHead();
              newNode := Result.AddAfter(newNode, newTok2);
              newTok3.isControl := false;
              newTok3.text := Copy(tok.Text, Length(tok.Text) - (count - LRange.head) + 1 -1, MaxInt); // +1?
              Result.AddAfter(newNode, newTok3);
              LNode.Free;
              break;
            end
          end;
          LNode := LNode.Next;
        end
      end;
      // process tail[i]
      count := 0;
      LNode := Result.First;
      while Assigned(LNode) do
      begin
        tok := LNode.Value;

        if (not tok.isControl) then
        begin
          Count := Count + Length(tok.Text);
          if (count{ - 1} = LRange.tail) then
          begin
            newTok.isControl := true;
            newTok.text := LRange.format.renderTail();
            Result.AddAfter(LNode, newTok);
            break;
          end else if (count{ - 1} > LRange.tail) then
          begin
            newTok1.isControl := false;
            newTok1.text := Copy(tok.Text, 1, Length(tok.Text) - (count - LRange.tail) + 1 -1);
            newNode := Result.AddAfter(LNode, newTok1);
            newTok2.isControl := true;
            newTok2.text := LRange.format.renderTail();
            newNode := Result.AddAfter(newNode, newTok2);
            newTok3.isControl := false;
            newTok3.text := Copy(tok.Text, Length(tok.Text) - (count - LRange.tail) + 1, MaxInt);
            Result.AddAfter(newNode, newTok3);
            LNode.Free;
            break;
          end
        end;
        LNode := LNode.Next;
      end
    end; // end for each char format


    // --------------------------------------------------
    // Insert footnote into token list.
    // --------------------------------------------------
    for i := 0 to FFootNotes.Count -1 do
    begin
      LPos := FFootNotes[i].Position;
      if (LPos >{=} FText.Length) then begin
        continue;
      end;

      count := 0;
      LNode := Result.First;
      while Assigned(LNode) do begin
        tok := LNode.Value;

        if ( not tok.isControl) then begin
          count := count + Length(tok.Text);
          if (count{ - 1}) = LPos then begin
            newTok.isControl := true;
            newTok.text := FFootnotes[i].render();
            Result.AddAfter(LNode, newTok);
            break;
          end else if (count{ - 1} > LPos) then begin
            newTok1.isControl := false;
            newTok1.text := Copy(tok.text, 1, Length(tok.Text) - (count - LPos) + 1 -1);
            newNode := Result.AddAfter(LNode, newTok1);

            newTok2.isControl := true;
            newTok2.text := FFootnotes[i].render();
            newNode := Result.AddAfter(newNode, newTok2);

            newTok3.isControl := false;
            newTok3.text := Copy(tok.text, Length(tok.Text) - (count - LPos) + 1, MaxInt);
            Result.AddAfter(newNode, newTok3);
            LNode.Free;
            break;
          end
        end;
        LNode := LNode.Next;
      end
    end;

    // --------------------------------------------------
    // Insert control words into token list.
    // --------------------------------------------------
    for i := 0 to FControlWords.Count -1 do
    begin
      LPos := FControlWords[i].Position;
      if (LPos >{=} FText.Length) then begin
        continue;
      end;

      count := 0;
      LNode := Result.First;
      while Assigned(LNode) do begin
        tok := LNode.Value;

        if (not tok.isControl) then begin
          count := count + Length(tok.Text);
          if (count {- 1} = LPos) then begin
            newTok.isControl := true;
            newTok.text := FControlWords[i].render();
            Result.AddAfter(LNode, newTok);
            break;
          end else if (count{ - 1} > LPos) then begin
            newTok1.isControl := false;
            newTok1.text := Copy(tok.text, 1, Length(tok.Text) - (count - LPos) + 1 -1);
            newNode := Result.AddAfter(LNode, newTok1);

            newTok2.isControl := true;
            newTok2.text := FControlWords[i].render();
            newNode := Result.AddAfter(newNode, newTok2);

            newTok3.isControl := false;
            newTok3.text := Copy(tok.text, Length(tok.Text) - (count - LPos) + 1);
            Result.AddAfter(newNode, newTok3);
            LNode.Free;
            break;
          end
        end;
        LNode := LNode.Next;
      end;
    end;


  finally
    LDRanges.Free;
  end;
end;

constructor TRtfParagraph.Create(AAllowFootnote, AAllowControlWord: Boolean);
begin
  inherited Create;
  FText := TStringBuilder.Create;
  FLinespacing := -1;
  FAlignment := alLeft; //Changed default to .Left as .None was spreading text accross page.
  FCharFormats := TRtfCharFormats.Create;
  FAllowFootnote := AAllowFootnote;
  FAllowControlWord := AAllowControlWord;
  FFootnotes := TRtfFootnotes.Create;
  FControlWords := TRtfFieldControlWords.Create;
  FBlockHead := '{\pard';
  FBlockTail := '\par}';
  FStartNewPage := False;
  FFirstLineIndent := 0;
end;

destructor TRtfParagraph.Destroy;
begin
  FreeAndNil(FText);
  FreeAndNil(FCharFormats);
  FreeAndNil(FFootnotes);
  FreeAndNil(FControlWords);
end;

function TRtfParagraph.ExtractTokenList(AList: TTokenList): string;
var
  LStringBuilder: TStringBuilder;
  LNode: TTokenList.TNode;
begin
  LStringBuilder := TStringBuilder.Create;
  try
    LNode := AList.First;
    while Assigned(LNode) do
    begin
      if (LNode.Value.isControl) then
      begin
        LStringBuilder.Append(LNode.Value.text);
      end
      else
      begin
        LStringBuilder.Append(TRtfUtility.unicodeEncode(LNode.Value.text));
      end;
      LNode := LNode.Next;
    end;
    Result := LStringBuilder.ToString;
  finally
    LStringBuilder.Free;
  end;
end;

function TRtfParagraph.GetDefaultCharFormat: TRtfCharFormat;
begin
  if not Assigned(FDefaultCharFormat) then
  begin
    FDefaultCharFormat := TRtfCharFormat.Create(-1, -1, FText.Length);
  end;
  Result := FDefaultCharFormat;
end;

function TRtfParagraph.Render: string;
var
  LStringBuilder: TStringBuilder;
  LTokenList: TTokenList;
begin
  LTokenList := BuildTokenList;
  LStringBuilder := TStringBuilder.Create(FBlockHead);
  try
    if (FStartNewPage) then begin
      LStringBuilder.Append('\pagebb');
    end;
    if (FLinespacing >= 0) then begin
      LStringBuilder.Append('\sl-' + IntToStr(TRtfUtility.pt2Twip(FLinespacing)) + '\slmult0');
    end;
    if (FMargins[drTop] > 0) then begin
      LStringBuilder.Append('\sb' + IntToStr(TRtfUtility.pt2Twip(FMargins[drTop])));
    end;
    if (FMargins[drBottom] > 0) then begin
      LStringBuilder.Append('\sa' + IntToStr(TRtfUtility.pt2Twip(FMargins[drBottom])));
    end;
    if (FMargins[drLeft] > 0) then begin
      LStringBuilder.Append('\li' + IntToStr(TRtfUtility.pt2Twip(FMargins[drLeft])));
    end;
    if (FMargins[drRight] > 0) then begin
      LStringBuilder.Append('\ri' + IntToStr(TRtfUtility.pt2Twip(FMargins[drRight])));
    end;
    if (FFirstLineIndent >= 0) then begin
      LStringBuilder.Append('\fi' + IntToStr(TRtfUtility.pt2Twip(FFirstLineIndent)));
    end;
    LStringBuilder.Append(AlignmentCode());
    LStringBuilder.AppendLine();

    // insert default char format intto the 1st position of _charFormats
    if Assigned(FDefaultCharFormat) then begin
      LStringBuilder.AppendLine(FDefaultCharFormat.renderHead());
    end;
    LStringBuilder.AppendLine(extractTokenList(LTokenList));
    if Assigned(FDefaultCharFormat) then begin
      LStringBuilder.Append(FDefaultCharFormat.renderTail());
    end;

    LStringBuilder.AppendLine(FBlockTail);
    Result := LStringBuilder.ToString;
  finally
    LStringBuilder.Free;
    LTokenList.Free;
  end;
end;

procedure TRtfParagraph.SetText(AText: string);
begin
  FText.Free;
  FText := TStringBuilder.Create(AText);
end;

{ TRtfParagraph.TDisjointRange }

constructor TRtfParagraph.TDisjointRange.Create(AHead, ATail: Integer; AFmt: TRtfCharFormat);
begin
  head := AHead;
  tail := ATail;
  Format := AFmt;
end;

{ TRtfTable }

constructor TRtfTable.Create(ARowCount, AColCount: Integer; AHorizontalWidth,
  AFontSize: Double);
var
  i,j: integer;
begin
  inherited Create;

  if (ARowCount < 1) or (AColCount < 1) then
  begin
    raise TRtfException.Create('The number of rows or columns is less than 1.');
  end;

  FFontSize := AFontSize;
  FAlignment := alNone;
  FRowCount := ARowCount;
  FColCount := AColCount;
  FRepresentativeList := TRtfTableCells.Create(false);
  FStartNewPage := False;
  FTitleRowCount := 0;

  FHeaderBackgroundColour := -1;
  FRowBackgroundColour := -1;
  FRowAltBackgroundColour := -1;
  // Set cell default width according to paper width
  FDefaultCellWidth := AHorizontalWidth / AColCount;

  SetLength(FCellPadding, FRowCount);
  SetLength(FCells, FRowCount);
  SetLength(FRowHeight, FRowCount);
  SetLength(FRowKeepInSamePage, FRowCount);
  for i := 0 to FRowCount-1 do
  begin
    FRowHeight[i] := 0;
    FRowKeepInSamePage[i] := false;
    FCellPadding[i] := TRtfMargins.Create;
    SetLength(FCells[i], FColCount);
    for j := 0 to FColCount-1 do
    begin
      FCells[i][j] := TRtfTableCell.Create(FDefaultCellWidth, i, j, Self);
    end;
  end;
end;

destructor TRtfTable.Destroy;
var
  i,j: integer;
begin
  FreeAndNil(FRepresentativeList);
  for i := 0 to FRowCount-1 do
  begin
    FCellPadding[i].Free;
    for j := 0 to FColCount-1 do
    begin
      FCells[i][j].Free;
    end;
  end;
  inherited;
end;

function TRtfTable.GetCell(ARow, ACol: integer): TRtfTableCell;
begin
  Result := nil;
  if (ARow >= 0) and (ARow < Length(FCells)) then
  begin
    if (ACol >= 0) and (ACol < Length(FCells[ARow])) then
    begin
      if (FCells[ARow][ACol].IsMerged) then begin
        Result := FCells[ARow][ACol].MergeInfo.Representative;
      end
      else
      begin
        Result := FCells[ARow, ACol];
      end;
    end;
  end;
  if not Assigned(Result) then
    raise TRtfException.Create('Cell does not exist (' + IntToStr(ARow) + '-'+ IntToStr(ACol) + '). ');

end;

function TRtfTable.GetCellPadding(index: integer): TRtfMargins;
begin
  Result := nil;
  if (index >= 0) and (index < Length(FCellPadding)) then
    Result := FCellPadding[index];
end;

function TRtfTable.GetDefaultCharFormat: TRtfCharFormat;
begin
  if not Assigned(FDefaultCharFormat) then
  begin
    FDefaultCharFormat := TRtfCharFormat.Create(-1, -1, 1);
  end;
  Result := FDefaultCharFormat;
end;

function TRtfTable.Merge(ATopRow, ALeftCol, ARowSpan,
  AColSpan: integer): TRtfTableCell;
var
  i,j: integer;
  LWidth: double;
begin
  if (ATopRow < 0) or (ATopRow >= FRowCount) then begin
    raise TRtfException.Create('Row index out of range');
  end;
  if (ALeftCol < 0) or (ALeftCol >= FColCount) then begin
    raise TRtfException.Create('Column index out of range');
  end;
  if (ARowSpan < 1) or (ATopRow + ARowSpan - 1 >= FRowCount) then begin
    raise TRtfException.Create('Row span out of range.');
  end;
  if (AColSpan < 1) or (ALeftCol + AColSpan - 1 >= FColCount) then begin
    raise TRtfException.Create('Column span out of range.');
  end;
  if (AColSpan = 1) and (ARowSpan = 1) then
  begin
    Result := Cell[ATopRow, ALeftCol];
    Exit;
  end;
  // Check if the cell has been merged before.
  for i := 0 to ARowSpan -1 do begin
    for j := 0 to AColSpan -1 do begin
      if (FCells[ATopRow + i][ALeftCol + j].IsMerged) then begin
        raise TRtfException.Create('Cannot merge cells because some of the cells has been merged.');
      end;
    end;
  end;

  LWidth := 0;
  for i := 0 to ARowSpan -1 do begin
    for j := 0 to AColSpan -1 do begin
      // Sum up the column widths in the first row.
      if (i = 0) then begin
        LWidth := LWidth + FCells[ATopRow][ALeftCol + j].Width;
      end;
      // Set merge info for each cell.
      // Note: The representatives of all cells are set to the (ATopRow, leftCol) cell.
      FCells[ATopRow + i][ALeftCol + j].MergeInfo := TRtfCellMergeInfo.Create(FCells[ATopRow][ALeftCol], ARowSpan, AColSpan, i, j);
      if (i <> 0) or (j <> 0) then begin
        // Transfer the blocks (contents) of each cell to their representative cell.
        FCells[ATopRow + i][ALeftCol + j].transferBlocksTo(FCells[ATopRow + i][ALeftCol + j].MergeInfo.Representative);
      end;
    end;
  end;

  // Set cell width in the representative cell.
  Result := FCells[ATopRow][ALeftCol];
  Result.Width := LWidth;
  FRepresentativeList.Add(Result);
end;

function TRtfTable.Render: string;
var
  LStringBuilder: TStringBuilder;
  i,j: integer;
  topMargin: double;
  colAcc, colRightPos: integer;
  nextCellLeftBorderClearance: double;
  d: TRtfDirection;
  bdr: TRtfBorder;
begin
  LStringBuilder := TStringBuilder.Create;
  try
    // validate borders for each cell.
    // (borders may be changed because of cell merging)
    validateAllMergedCellBorders();
    // set default char format for each cell.
    if Assigned(FDefaultCharFormat) then
    begin
      for i := 0 to FRowCount -1 do
      begin
        for j := 0 to FColCount -1 do
        begin
          if FCells[i][j].IsMerged and (FCells[i][j].MergeInfo.Representative <> FCells[i][j]) then begin
            continue;
          end;
          if Assigned(FCells[i][j].DefaultCharFormat) then begin
            FCells[i][j].DefaultCharFormat.copyFrom(FDefaultCharFormat);
          end
        end
      end
    end;

    topMargin := FMargins[drTop] - FFontSize;
    if FStartNewPage or (topMargin > 0) then begin
      LStringBuilder.Append('{\pard');
      if (FStartNewPage) then begin
        LStringBuilder.Append('\pagebb');
      end;
      if (FMargins[drTop] >= 0) then begin
        LStringBuilder.Append('\sl-' + IntToStr(TRtfUtility.pt2Twip(topMargin)));
      end else begin
        LStringBuilder.Append('\sl-1');
      end;
      LStringBuilder.AppendLine('\slmult0\par}');
    end;

    for i := 0 to FRowCount -1 do
    begin
      colAcc := 0;
      LStringBuilder.Append('{\trowd\trgaph' +
                    Format('\trpaddl%d\trpaddt%d\trpaddr%d\trpaddb%d',
                                  [TRtfUtility.pt2Twip(CellPadding[i][drLeft]),
                                  TRtfUtility.pt2Twip(CellPadding[i][drTop]),
                                  TRtfUtility.pt2Twip(CellPadding[i][drRight]),
                                  TRtfUtility.pt2Twip(CellPadding[i][drBottom])]));
      case (FAlignment) of
        alLeft:           LStringBuilder.Append('\trql');
        alRight:          LStringBuilder.Append('\trqr');
        alCenter:         LStringBuilder.Append('\trqc');
        alFullyJustified: LStringBuilder.Append('\trqj');
      end;
      LStringBuilder.AppendLine();
      if (FMargins[drLeft] >= 0) then begin
        LStringBuilder.AppendLine('\trleft' + IntToStr(TRtfUtility.pt2Twip(FMargins[drLeft])));
        colAcc := TRtfUtility.pt2Twip(FMargins[drLeft]);
      end;
      if (FRowHeight[i] > 0) then begin
        LStringBuilder.Append('\trrh' + IntToStr(TRtfUtility.pt2Twip(FRowHeight[i])));
      end;
      if (FRowKeepInSamePage[i]) then begin
        LStringBuilder.Append('\trkeep');
      end;
      if (i < FTitleRowCount) then begin
        LStringBuilder.Append('\trhdr');
      end;
      LStringBuilder.AppendLine();

      for j := 0 to FColCount -1 do
      begin
        if (FCells[i][j].IsMerged and not FCells[i][j].IsBeginOfColSpan) then begin
          continue;
        end;
        if j < FColCount - 1 then
          nextCellLeftBorderClearance := cell[i, j + 1].OuterLeftBorderClearance
        else
          nextCellLeftBorderClearance := 0;
        colAcc := colAcc + TRtfUtility.pt2Twip(cell[i, j].Width);
        colRightPos := colAcc;
        if(nextCellLeftBorderClearance < 0) then begin
          colRightPos := colRightPos + TRtfUtility.pt2Twip(nextCellLeftBorderClearance);
          if colRightPos = 0 then
            colRightPos := 1;
        end;

        // Borders
        for d := Low(TRtfDirection) to High(TRtfDirection) do begin
          bdr := cell[i, j].Borders[d];
          if (bdr.Style <> bsNone) then begin
            LStringBuilder.Append('\clbrdr');
            case d of
              drTop:    LStringBuilder.Append('t');
              drRight:  LStringBuilder.Append('r');
              drBottom: LStringBuilder.Append('b');
              drLeft:   LStringBuilder.Append('l');
            end;
            LStringBuilder.Append('\brdrw' + IntToStr(TRtfUtility.pt2Twip(bdr.Width)));
            LStringBuilder.Append('\brdr');
            case bdr.Style of
              bsSingle: LStringBuilder.Append('s');
              bsDotted: LStringBuilder.Append('dot');
              bsDashed: LStringBuilder.Append('dash');
              bsDouble: LStringBuilder.Append('db');
            else
              raise TRtfException.Create('Unkown border style');
            end;
            LStringBuilder.Append('\brdrcf' + IntToStr(bdr.ColorDesc));
          end;
        end;

        // Cell background colour
        if (cell[i, j].BackgroundColorIndex <> -1) then
        begin
          LStringBuilder.Append(Format('\clcbpat%d', [cell[i, j].BackgroundColorIndex])); // cell.BackGroundColor overrides others
        end else if (i = 0) and (HeaderBackgroundColour <> -1) then begin
          LStringBuilder.Append(Format('\clcbpat%d', [HeaderBackgroundColour])); // header
        end else if (RowBackgroundColour <> -1) and ((RowAltBackgroundColour = -1) or (i mod 2 = 0)) then begin
         LStringBuilder.Append(Format('\clcbpat%d', [RowBackgroundColour])); // row colour
        end else if (RowBackgroundColour <> -1) and (RowAltBackgroundColour <> -1) and (i mod 2 <> 0) then begin
          LStringBuilder.Append(Format('\clcbpat%d', [RowAltBackgroundColour])); // alt row colour
        end;

        if FCells[i][j].IsMerged and (FCells[i][j].MergeInfo.RowSpan > 1) then begin
            if (FCells[i][j].IsBeginOfRowSpan) then begin
                LStringBuilder.Append('\clvmgf');
            end else begin
                LStringBuilder.Append('\clvmrg');
            end
        end;
        case (FCells[i][j].VAlign) of
            avTop:    LStringBuilder.Append('\clvertalt');
            avMiddle: LStringBuilder.Append('\clvertalc');
            avBottom: LStringBuilder.Append('\clvertalb');
        end;
        LStringBuilder.AppendLine('\cellx' + IntToStr(colRightPos));
      end;

      for j := 0 to FColCount -1 do
      begin
        if (not FCells[i][j].IsMerged) or FCells[i][j].IsBeginOfColSpan then begin
          LStringBuilder.Append(FCells[i][j].render());
        end;
      end;

      LStringBuilder.AppendLine('\row}');
    end;

    if (FMargins[drBottom] >= 0) then begin
      LStringBuilder.Append('\sl-' + IntToStr(TRtfUtility.pt2Twip(FMargins[drBottom])) + '\slmult');
    end;

    Result := LStringBuilder.ToString();
  finally
    LStringBuilder.Free;
  end;
end;

procedure TRtfTable.SetColWidth(ACol: integer; AWidth: Double);
var
  i: integer;
begin
  if (ACol < 0) or (ACol >= FColCount) then
  begin
    raise TRtfException.Create('Column index out of range');
  end;
  for i := 0 to FRowCount -1 do
  begin
    if (FCells[i][ACol].IsMerged) then begin
      raise TRtfException.Create('Column width cannot be set because some cell in this column has been merged.');
    end;
  end;
  for i := 0 to FRowCount -1 do
  begin
    FCells[i][ACol].Width := AWidth;
  end;
end;

procedure TRtfTable.SetHeaderBorderColors(AColorOuter,
  AColorInner: TRtfColorDescriptor);
var
  j: integer;
begin
  for j := 0 to FcolCount -1 do
  begin
    FCells[0][j].Borders[drTop].ColorDesc := AColorOuter;
    FCells[0][j].Borders[drBottom].ColorDesc := AColorInner;
    if (j = 0) then
    begin
      // The first column
      FCells[0][j].Borders[drRight].ColorDesc := AColorInner;
      FCells[0][j].Borders[drLeft].ColorDesc := AColorOuter;
    end
    else if (j = (FColCount - 1)) then
    begin
      // The last column
      FCells[0][j].Borders[drRight].ColorDesc := AColorOuter;
      FCells[0][j].Borders[drLeft].ColorDesc := AColorInner;
    end
    else
    begin
      FCells[0][j].Borders[drRight].ColorDesc := AColorInner;
      FCells[0][j].Borders[drLeft].ColorDesc := AColorInner;
    end;
  end;
end;

procedure TRtfTable.SetInnerBorder(AStyle: TRtfBorderStyle; AWidth: double);
begin
  SetInnerBorder(AStyle, AWidth, 0);
end;

procedure TRtfTable.SetInnerBorder(AStyle: TRtfBorderStyle; AWidth: double;
  AColor: TRtfColorDescriptor);
var
  i,j: integer;
begin
  for i := 0 to FRowCount -1 do
  begin
    for j := 0 to FcolCount -1 do
    begin
      if (i = 0) then begin
        // The first row
        FCells[i][j].Borders[drBottom].Style := AStyle;
        FCells[i][j].Borders[drBottom].Width := AWidth;
        FCells[i][j].Borders[drBottom].ColorDesc := AColor;
      end else if (i = FRowCount - 1) then begin
        // The last row
        FCells[i][j].Borders[drTop].Style := AStyle;
        FCells[i][j].Borders[drTop].Width := AWidth;
        FCells[i][j].Borders[drTop].ColorDesc := AColor;
      end else begin
        FCells[i][j].Borders[drTop].Style := AStyle;
        FCells[i][j].Borders[drTop].Width := AWidth;
        FCells[i][j].Borders[drTop].ColorDesc := AColor;
        FCells[i][j].Borders[drBottom].Style := AStyle;
        FCells[i][j].Borders[drBottom].ColorDesc := AColor;
        FCells[i][j].Borders[drBottom].Width := AWidth;
      end;
      if (j = 0) then begin
        // The first column
        FCells[i][j].Borders[drRight].Style := AStyle;
        FCells[i][j].Borders[drRight].Width := AWidth;
        FCells[i][j].Borders[drRight].ColorDesc := AColor;
      end else if (j = FColCount - 1) then begin
        // The last column
        FCells[i][j].Borders[drLeft].Style := AStyle;
        FCells[i][j].Borders[drLeft].Width := AWidth;
        FCells[i][j].Borders[drLeft].ColorDesc := AColor;
      end else begin
        FCells[i][j].Borders[drRight].Style := AStyle;
        FCells[i][j].Borders[drRight].Width := AWidth;
        FCells[i][j].Borders[drRight].ColorDesc := AColor;
        FCells[i][j].Borders[drLeft].Style := AStyle;
        FCells[i][j].Borders[drLeft].Width := AWidth;
        FCells[i][j].Borders[drLeft].ColorDesc := AColor;
      end;
    end;
  end;
end;

procedure TRtfTable.SetOuterBorder(AStyle: TRtfBorderStyle; AWidth: double);
begin
  SetOuterBorder(AStyle, AWidth, 0);
end;

procedure TRtfTable.SetOuterBorder(AStyle: TRtfBorderStyle; AWidth: double;
  AColor: TRtfColorDescriptor);
var
  i: integer;
begin
  for i := 0 to FColCount -1 do
  begin
    FCells[0][i].Borders[drTop].Style := AStyle;
    FCells[0][i].Borders[drTop].Width := AWidth;
    FCells[0][i].Borders[drTop].ColorDesc := AColor;
    FCells[FRowCount - 1][i].Borders[drBottom].Style := AStyle;
    FCells[FRowCount - 1][i].Borders[drBottom].Width := AWidth;
    FCells[FRowCount - 1][i].Borders[drBottom].ColorDesc := AColor;
  end;
  for i := 0 to FRowCount -1 do
  begin
    FCells[i][0].Borders[drLeft].Style := AStyle;
    FCells[i][0].Borders[drLeft].Width := AWidth;
    FCells[i][0].Borders[drLeft].ColorDesc := AColor;
    FCells[i][FColCount - 1].Borders[drRight].Style := AStyle;
    FCells[i][FColCount - 1].Borders[drRight].Width := AWidth;
    FCells[i][FColCount - 1].Borders[drRight].ColorDesc := AColor;
  end;
end;

procedure TRtfTable.SetRowHeight(ARow: integer; AHeight: Double);
var
  i: integer;
begin
  if (ARow < 0) or (ARow >= FRowCount) then
  begin
    raise TRtfException.Create('Row index out of range');
  end;
  for i := 0 to FColCount -1 do
  begin
    if (FCells[ARow][i].IsMerged) then begin
      raise TRtfException.Create('Row height cannot be set because some cell in this row has been merged.');
    end;
  end;
  FRowHeight[ARow] := AHeight;
end;

procedure TRtfTable.SetRowKeepInSamePage(ARow: integer; AAllow: boolean);
begin
  if (ARow < 0) or (ARow >= FRowCount) then
  begin
    raise TRtfException.Create('Row index out of range');
  end;
  FRowKeepInSamePage[ARow] := AAllow;
end;

procedure TRtfTable.ValidateAllMergedCellBorders;
var
  i: Integer;
begin
  for i := 0 to FRepresentativeList.Count -1 do
  begin
    ValidateMergedCellBorders(FRepresentativeList[i]);
  end;
end;

procedure TRtfTable.ValidateMergedCellBorder(ARepresentative: TRtfTableCell;
  ADir: TRtfDirection);
var
  stat: TDictionary<TRtfBorder,integer>;
  de: TPair<TRtfBorder, integer>;
  majorityBorder: TRtfBorder;
  majorityCount, limit, i: integer;
  r,c: integer;
  bdr: TRtfBorder;
begin
  if (not ARepresentative.IsMerged) then begin
    raise TRtfException.Create('Invalid representative (cell is not merged).');
  end;

  stat := TDictionary<TRtfBorder,integer>.Create;
  try
    if (ADir = drTop) or (ADir = drBottom) then
    begin
      limit := ARepresentative.MergeInfo.ColSpan;
    end
    else
    begin
      limit := ARepresentative.MergeInfo.RowSpan;
    end;

    for i := 0 to limit -1 do begin
      if (ADir = drTop) or (ADir = drBottom) then begin
        if (ADir = drTop) then begin
          r := 0;
        end else begin // ADir == bottom
          r := ARepresentative.MergeInfo.RowSpan - 1;
        end;
        c := i;
      end else begin // ADir == right || left
        if (ADir = drRight) then begin
          c := ARepresentative.MergeInfo.ColSpan - 1;
        end else begin // ADir == left
          c := 0;
        end;
        r := i;
      end;
      bdr := FCells[ARepresentative.RowIndex + r][ARepresentative.ColIndex + c].Borders[ADir];
      if (stat.ContainsKey(bdr)) then begin
        stat[bdr] := stat[bdr] + 1;
      end else begin
        stat.Add(bdr, 1);
      end;
    end;
    majorityCount := -1;
    majorityBorder := ARepresentative.Borders[ADir];
    for de in stat do
    begin
      if(de.Value > majorityCount) then
      begin
        majorityCount := de.Value;
        majorityBorder.Style := de.Key.Style;
        majorityBorder.Width := de.Key.Width;
        majorityBorder.ColorDesc := de.Key.ColorDesc;
      end;
    end;
  finally
    stat.Free;
  end;
end;

procedure TRtfTable.ValidateMergedCellBorders(ARepresentative: TRtfTableCell);
begin
  if (not ARepresentative.IsMerged) then begin
    raise TRtfException.Create('Invalid representative (cell is not merged).');
  end;
  ValidateMergedCellBorder(ARepresentative, drTop);
  ValidateMergedCellBorder(ARepresentative, drRight);
  ValidateMergedCellBorder(ARepresentative, drBottom);
  ValidateMergedCellBorder(ARepresentative, drLeft);
end;

{ TRtfTableCell }

constructor TRtfTableCell.Create(AWidth: double; ARowIndex, AColIndex: integer;
  AParentTable: TRtfTable);
begin
  inherited Create(True, False);
  FWidth := AWidth;
  FHAlign := TRtfAlign.alNone;
  FVAlign := TRtfAlignVertical.avTop;
  FBorders := TRtfBorders.Create;
  FMergeInfo := nil;
  FRowIndex := ARowIndex;
  FColIndex := AColIndex;
  FBackgroundColorIndex := -1;
  FParentTable := AParentTable;
end;

destructor TRtfTableCell.Destroy;
begin
  FreeAndNil(FBorders);
  FreeAndNil(FMergeInfo);
  inherited;
end;

function TRtfTableCell.IsBeginOfColSpan: boolean;
begin
  Result := Assigned(FMergeInfo) and (FMergeInfo.ColIndex = 0);
end;

function TRtfTableCell.IsBeginOfRowSpan: boolean;
begin
  Result := Assigned(FMergeInfo) and (FMergeInfo.RowIndex = 0);
end;

function TRtfTableCell.IsMerged: boolean;
begin
  Result := Assigned(FMergeInfo);
end;

function TRtfTableCell.Render: string;
var
  LStringBuilder: TStringBuilder;
  LAlign: string;
  i: integer;
  LBlock: TRtfBlock;
begin
  LStringBuilder := TStringBuilder.Create;
  try
    case FHAlign of
      alLeft: LAlign := '\ql';
      alRight: LAlign := '\qr';
      alCenter: LAlign := '\qc';
      alFullyJustified: LAlign := '\qj';
      alDistributed: LAlign := '\qd';
    else
      LAlign := '';
    end;

    if (FBlocks.Count <= 0) then begin
      LStringBuilder.AppendLine('\pard\intbl');
    end else begin
      for i := 0 to FBlocks.Count -1 do begin
        LBlock := FBlocks[i];
        if Assigned(FDefaultCharFormat) and Assigned(LBlock.DefaultCharFormat) then begin
          LBlock.DefaultCharFormat.copyFrom(FDefaultCharFormat);
        end;
        if (LBlock.Margins[drTop] < 0) then begin
          LBlock.Margins[drTop] := 0;
        end;
        if (LBlock.Margins[drRight] < 0) then begin
          LBlock.Margins[drRight] := 0;
        end;
        if (LBlock.Margins[drBottom] < 0) then begin
          LBlock.Margins[drBottom] := 0;
        end;
        if (LBlock.Margins[drLeft] < 0) then begin
          LBlock.Margins[drLeft] := 0;
        end;
        if (i = 0) then begin
          LBlock.BlockHead := '\pard\intbl' + LAlign;
        end else begin
          LBlock.BlockHead := '\par' + LAlign;
        end;
        LBlock.BlockTail := '';
        LStringBuilder.AppendLine(LBlock.render());
      end;
    end;

    LStringBuilder.AppendLine('\cell');
    Result := LStringBuilder.ToString();
  finally
    LStringBuilder.Free;
  end;
end;

procedure TRtfTableCell.SetBorderColor(AColor: TRtfColorDescriptor);
begin
  FBorders[drTop].ColorDesc := AColor;
  FBorders[drBottom].ColorDesc := AColor;
  FBorders[drLeft].ColorDesc := AColor;
  FBorders[drRight].ColorDesc := AColor;
end;

{ TRtfImage }

constructor TRtfImage.Create(AFileName: string);
var
  LExt: string;
  grpCls: TGraphicClass;
  grp: TGraphic;
begin
  LExt := ExtractFileExt(AFileName);
  grpCls := FindGraphicClass(LExt);
  if Assigned(grpCls) then
  begin
    grp := grpCls.Create;
    try
      grp.LoadFromFile(AFileName);
      Create(grp);
    finally
      grp.Free;
    end;
  end else
    raise TRtfException.Create('File format not supported: ' + AFileName);
end;

constructor TRtfImage.Create(AImageStream: TStream);
var
  grpCls: TGraphicClass;
  grp: TGraphic;
begin
  grpCls := FindGraphicClass(AImageStream);
  if Assigned(grpCls) then
  begin
    grp := grpCls.Create;
    try
      grp.LoadFromStream(AImageStream);
      Create(grp);
    finally
      grp.Free;
    end;
  end else
    raise TRtfException.Create('File format not supported.');
end;

procedure TRtfImage.BitmapToDIBStream(ABitmap: TBitmap; AStream: TStream);
var
  LHeader, LBits: PAnsiChar;
  HeaderSize, LBitsSize: DWORD;
begin
  FBmpWidth := ABitmap.Width;
  FBmpHeight := ABitmap.Height;
  GetDIBSizes(ABitmap.Handle, HeaderSize, LBitsSize);
  GetMem(LHeader, HeaderSize + LBitsSize);
  try
    LBits := LHeader + HeaderSize;
    GetDIB(ABitmap.Handle, ABitmap.Palette, LHeader^, LBits^);
    AStream.Write(LHeader^, (HeaderSize + LBitsSize));
  finally
    FreeMem(LHeader);
  end;
end;

constructor TRtfImage.Create(AGraphic: TGraphic);
var
  LMemStream: TMemoryStream;
begin
  inherited Create;
  if not Assigned(AGraphic) then
    raise TRtfException.Create('No graphic for image');
  if AGraphic is TJPEGImage then begin
    FImgType := itJpg;
  end
  else if AGraphic is TPngImage then begin
    FImgType := itPng;
  end
  else if AGraphic is TGifImage then begin
    FImgType := itGif;
  end
  else if AGraphic is TBitmap then begin
    FImgType := itBmp;
  end
  else
  begin
    raise TRtfException.Create('File format not supported: ' + AGraphic.ClassName);
  end;

  FBmpWidth := 0;
  FBmpHeight := 0;
  FKeepAspectRatio := True;

  LMemStream := TMemoryStream.Create;
  try
    if AGraphic is TBitmap then
    begin
      BitmapToDIBStream(AGraphic as TBitmap, LMemStream);
    end
    else
    begin
      AGraphic.SaveToStream(LMemStream);
    end;
    LMemStream.Position := 0;
    SetLength(FImgBytes, LMemStream.Size);
    LMemStream.ReadBuffer(FImgBytes[0], LMemStream.Size);
  finally
    LMemStream.Free;
  end;
  FWidth := AGraphic.Width;
  FHeight := AGraphic.Height;
  Init;
end;

destructor TRtfImage.Destroy;
begin

  inherited;
end;

function TRtfImage.ExtractImage: string;
var
  LStringBuilder: TStringBuilder;
  i: integer;
begin
  LStringBuilder := TStringBuilder.Create;
  try
    for i := 0 to Length(FImgBytes) -1 do
    begin
      if (i <> 0) and (i mod 60 = 0) then
      begin
        LStringBuilder.AppendLine();
      end;
      LStringBuilder.Append(IntToHex(FImgBytes[i],2));
    end;
    Result := LStringBuilder.ToString();
  finally
    LStringBuilder.Free;
  end;
end;

const
  MinGraphicSize = 44; //we may test up to & including the 11th longword

function TRtfImage.FindGraphicClass(AStream: TStream): TGraphicClass;
var
  LBuffer: PByte;
  LCurPos: Int64;
  LBytesRead: Integer;
begin
  if AStream is TCustomMemoryStream then begin
    LBuffer := TCustomMemoryStream(AStream).Memory;
    LCurPos := AStream.Position;
    Inc(LBuffer, LCurPos);
    Result := FindGraphicClass(LBuffer^, AStream.Size - LCurPos);
    Exit;
  end;

  GetMem(LBuffer, MinGraphicSize);
  try
    LBytesRead := AStream.Read(LBuffer^, MinGraphicSize);
    AStream.Seek(-LBytesRead, soCurrent);
    Result := FindGraphicClass(LBuffer^, LBytesRead);
  finally
    FreeMem(LBuffer);
  end;
end;

function TRtfImage.FindGraphicClass(const ABuffer;
  const ABufferSize: Int64): TGraphicClass;
var
  LongWords: array[Byte] of LongWord absolute ABuffer;
  Words: array[Byte] of Word absolute ABuffer;
begin
  Result := nil;
  if ABufferSize < MinGraphicSize then Exit;
  case Words[0] of
    $4D42: Result := TBitmap;
    $D8FF: Result := TJPEGImage;
  else
    if Int64(ABuffer) = $A1A0A0D474E5089 then
      Result := TPngImage
//    else if StrLComp(PAnsiChar(@Buffer), '<svg', 4) = 0 then
//      Result := TSVGImageFile
    else if LongWords[0] = $9AC6CDD7 then
      Result := TMetafile
    else if (LongWords[0] = 1) and (LongWords[10] = $464D4520) then
      Result := TMetafile
    else if StrLComp(PAnsiChar(@ABuffer), 'GIF', 3) = 0 then
      Result := TGIFImage
    else if Words[1] = 1 then
      Result := TIcon;
  end;
end;

function TRtfImage.FindGraphicClass(AExt: string): TGraphicClass;
begin
  if CompareText(AExt, '.bmp') = 0 then
    Result := TBitmap
  else if CompareText(AExt, '.png') = 0 then
    Result := TPngImage
  else if CompareText(AExt, '.jpg') = 0 then
    Result := TJPEGImage
  else if CompareText(AExt, '.jpeg') = 0 then
    Result := TJPEGImage
  else if CompareText(AExt, '.wmf') = 0 then
    Result := TMetafile
  else if CompareText(AExt, '.emf') = 0 then
    Result := TMetafile
  else
    Result := nil;
end;

procedure TRtfImage.Init;
begin
  FKeepAspectRatio := true;
  FBlockHead := '{\pard';
  FBlockTail := '}';
  FStartNewPara := False;
  FStartNewPage := False;
end;

function TRtfImage.Render: string;
var
  LStringBuilder: TStringBuilder;
begin
  LStringBuilder := TStringBuilder.Create(FBlockHead);
  try
    if (FStartNewPage) then begin
      LStringBuilder.Append('\pagebb');
    end;

    if (FMargins[drTop] >= 0) then begin
      LStringBuilder.Append('\sb' + IntToStr(TRtfUtility.pt2Twip(FMargins[drTop])));
    end;
    if (FMargins[drBottom] >= 0) then begin
      LStringBuilder.Append('\sa' + IntToStr(TRtfUtility.pt2Twip(FMargins[drBottom])));
    end;
    if (FMargins[drLeft] >= 0) then begin
      LStringBuilder.Append('\li' + IntToStr(TRtfUtility.pt2Twip(FMargins[drLeft])));
    end;
    if (FMargins[drRight] >= 0) then begin
      LStringBuilder.Append('\ri' + IntToStr(TRtfUtility.pt2Twip(FMargins[drRight])));
    end;
    case (FAlignment) of
      alLeft:   LStringBuilder.Append('\ql');
      alRight:  LStringBuilder.Append('\qr');
      alCenter: LStringBuilder.Append('\qc');
    end;
    LStringBuilder.AppendLine();

    if (FImgType in [itJpg, itGif, itPng]) then begin
      LStringBuilder.Append('{\*\shppict');
    end;
    LStringBuilder.Append('{\pict');
    if (FImgType = itJpg) then begin
      LStringBuilder.Append('\jpegblip');
    end else if (FImgType = itPng) or (FImgType = itGif) then begin
      LStringBuilder.Append('\pngblip');
    end else if (FImgType = itBmp) then begin
      LStringBuilder.AppendFormat('\picw%d\pich%d\dibitmap0', [FBmpWidth, FBmpHeight]);
    end else begin
      raise TRtfException.Create('Image type not supported.');
    end;
    if (FHeight > 0) then begin
      LStringBuilder.Append('\pichgoal' + IntToStr(TRtfUtility.pt2Twip(FHeight)));
    end;
    if (FWidth > 0) then begin
      LStringBuilder.Append('\picwgoal' + IntToStr(TRtfUtility.pt2Twip(FWidth)));
    end;
    LStringBuilder.AppendLine();

    LStringBuilder.AppendLine(extractImage());
    if (FImgType in [itJpg, itGif, itPng]) then begin
      LStringBuilder.Append('}');
    end;
    LStringBuilder.AppendLine('}');
    if (FStartNewPara) then
      LStringBuilder.Append('\par');
    LStringBuilder.AppendLine(FBlockTail);

    Result := LStringBuilder.ToString;
  finally
    LStringBuilder.Free;
  end;
end;

procedure TRtfImage.SetHeight(const Value: double);
var
  ratio: double;
begin
  if FKeepAspectRatio and (FHeight > 0) then
  begin
    ratio := FWidth / FHeight;
    FWidth := Value * ratio;
  end;
  FHeight := Value;
end;

procedure TRtfImage.SetWidth(const Value: double);
var
  ratio: double;
begin
  if FKeepAspectRatio and (FWidth > 0) then
  begin
    ratio := FHeight / FWidth;
    FHeight := Value * ratio;
  end;
  FWidth := Value;
end;

{ TRtfSectionFooter }

constructor TRtfSectionFooter.Create(AParent: TRtfSection);
begin
  inherited Create(True, True, True, True, True, True, False);
  if not Assigned(AParent) then
    raise TRtfException.Create('Section footer can only be placed within a section.');
end;

function TRtfSectionFooter.Render: string;
var
  LStringBuilder: TStringBuilder;
begin
  LStringBuilder := TStringBuilder.Create;
  try
    LStringBuilder.AppendLine('{\footerr \ltrpar \pard\plain');
    LStringBuilder.AppendLine('\par ');
    LStringBuilder.Append( inherited Render);
    LStringBuilder.AppendLine('\par');
    LStringBuilder.AppendLine('}');

    Result := LStringBuilder.ToString();
  finally
    LStringBuilder.Free;
  end;
end;

{ TRtfSection }

constructor TRtfSection.Create(AType: TRtfSectionStartsEnd; ADoc: TRtfDocument);
begin
  inherited Create;
  FParentDocument := ADoc;
  FAlignment := alNone;
  FPageOrientation := poPortrait;
  FStartEnd := AType;
  FFooterPositionFromPageBottom := 720;
  FSectionFooter := nil;
end;

destructor TRtfSection.Destroy;
begin
  FreeAndNil(FSectionFooter);
  inherited;
end;

function TRtfSection.GetDefaultCharFormat: TRtfCharFormat;
begin
  raise TRtfException.Create('DefaultCharFormat is not supported for sections.');
end;

function TRtfSection.GetSectionFooter: TRtfSectionFooter;
begin
  if not Assigned(FSectionFooter) then
  begin
    FSectionFooter := TRtfSectionFooter.Create(Self);
  end;
  Result := FSectionFooter;
end;

function TRtfSection.Render: string;
var
  LStringBuilder: TStringBuilder;
begin
  LStringBuilder := TStringBuilder.Create;
  try
    if (StartEnd = sseStart) then
    begin
      LStringBuilder.AppendLine(Format('{{\sectd\ltrsect\footery%d\sectdefaultcl\sftnbj%d ', [FooterPositionFromPageBottom, AlignmentCode]));
      if (PageOrientation = poLandscape) then
      begin
        LStringBuilder.Append('\lndscpsxn ');
      end;
      LStringBuilder.Append(Format('\pgwsxn%d\pghsxn%d ', [PageWidth, PageHeight]));
      if ( not ParentDocument.Margins.Equals(Margins)) then
      begin
        LStringBuilder.Append(Format('\marglsxn%d\margrsxn%d\margtsxn%d\margbsxn%d ',
                                      [Margins[drLeft], Margins[drRight], Margins[drTop], Margins[drBottom]]));
      end;
      if Assigned(FSectionFooter) then
      begin
        LStringBuilder.AppendLine( SectionFooter.render );
      end;
    end
    else
    begin
      LStringBuilder.AppendLine('\sect }}');
    end;
    Result := LStringBuilder.ToString();
  finally
    LStringBuilder.Free;
  end;
end;

{ TRtfFootNote }

constructor TRtfFootNote.Create(APosition, ATextLength: integer);
begin
  inherited Create;
  if (APosition < 1) or (APosition > ATextLength) then
    raise TRtfException.CreateFmt('Invalid footnote position: %d (text length=%d)', [APosition, ATextLength]);
  FPosition := APosition;
end;

function TRtfFootNote.Render: string;
var
  LStringBuilder: TStringBuilder;
begin
  LStringBuilder := TStringBuilder.Create;
  try
    LStringBuilder.AppendLine('{\super\chftn}');
    LStringBuilder.AppendLine('{\footnote\plain\chftn');
    FBlocks[FBlocks.Count - 1].BlockTail := '}';
    LStringBuilder.Append(inherited Render() );
    LStringBuilder.AppendLine('}');
    Result := LStringBuilder.ToString();
  finally
    LStringBuilder.Free;
  end;
end;

{ TRtfFieldControlWord }

constructor TRtfFieldControlWord.Create(APosition: integer; AType: TFieldType);
begin
  inherited Create;
  FPosition := APosition;
  FType := AType;
end;

function TRtfFieldControlWord.Render: string;
const
  CONTROLWORDPOOL: array[TFieldType] of string = (
    '',
    '{\field{\*\fldinst PAGE }}',
    '{\field{\*\fldinst NUMPAGES }}',
    '{\field{\*\fldinst DATE }}',
    '{\field{\*\fldinst TIME }}');
begin
  Result := CONTROLWORDPOOL[FType];
end;

{ TRtfSpecialInsert }

constructor TRtfSpecialInsert.Create;
begin
  inherited;
  FBlockHead := '{\pard';
  FBlockTail := '\par}';
  FStartNewPage := False;
end;

function TRtfSpecialInsert.Render: string;
begin
  Result := FBlockHead + Special + FBlockTail;
end;

function TRtfSpecialInsert.Special: string;
begin
  Result := '';
end;

{ TRtfSpecialInsertHR }

function TRtfSpecialInsertHR.Special: string;
begin
  Result := '\brdrb \brdrs\brdrw10\brsp20 {\fs4\~}';
end;

{ TRtfListItem }

destructor TRtfListItem.Destroy;
begin
  FreeAndNil(FParagraph);
  inherited;
end;

function TRtfListItem.GetParagraph: TRtfParagraph;
begin
  if not Assigned(FParagraph) then
  begin
    FParagraph := TRtfParagraph.Create;
    FParagraph.BlockTail := '';
    FParagraph.BlockHead := '';
    FParagraph.FirstLineindent := -1;
  end;
  Result := FParagraph;
end;

function TRtfListItem.Render: string;
begin
  Result := '';
  if Assigned(FParagraph) then
    Result := FParagraph.Render;
end;

{ TRtfListItems }

function TRtfListItems.AddList(AType: TRtfListType): TRtfListItems;
begin
  Result := TRtfListItems.Create(Self, AType, FDocument);
  FItems.Add(Result);
end;

constructor TRtfListItems.Create(AParent: TRtfListItems; AType: TRtfListType; ADoc: TRtfDocument);
begin
  inherited Create;
  FParent := AParent;
  FDocument := ADoc;
  FTextindent := 40;
  FListIndent := 40;
  FItems := TObjectList<TRtfListItem>.Create;
  ListType := AType;
  FStartNumber := 1;
end;

destructor TRtfListItems.Destroy;
begin
  FreeAndNil(FItems);
  inherited;
end;

function TRtfListItems.AddItem: TRtfListItem;
begin
  Result := TRtfListItem.Create;
  FItems.Add(Result);
end;

function TRtfListItems.GetAlphaNumber(ANumber: Integer;
  ALowercase: Boolean): string;
var
  LModulus: Integer;
begin
  Result := '';
  while (ANumber > 0) do
  begin
    LModulus := (ANumber - 1) mod 26;
    Result := Char(LModulus + 65) + Result;
    ANumber := (ANumber - 1) div 26;
  end;
  if ALowercase then
    Result := lowercase(Result);
end;

function TRtfListItems.GetDefaultCharFormat: TRtfCharFormat;
begin
  if not Assigned(FDefaultCharFormat) then
  begin
    FDefaultCharFormat := TRtfCharFormat.Create(-1, -1, 0);
  end;
  Result := FDefaultCharFormat;
end;

function TRtfListItems.GetListCharacter(ANumber: Integer): string;
begin
  case FListType of
    ltArrow:   Result := '\''d8';
    ltCircle:  Result := ' o';
    ltSquare:  Result := ' v';
    ltDiamond: Result := '\''a7';
    ltBullet:  Result := '\''B7';
  else
    Result := FNumberPrefix + GetNumber(ANumber) + FNumberSuffix;
  end;
end;

function TRtfListItems.GetListCharFormat: TRtfCharFormat;
begin
  if Assigned(FDefaultCharFormat) then
    Result := FDefaultCharFormat
  else if Assigned(FParent) and Assigned(FParent.FDefaultCharFormat) and
    not (FListType in [ltArrow, ltCircle, ltSquare, ltDiamond, ltBullet]) then
    Result := FParent.FDefaultCharFormat
  else
    Result := FDocument.DefaultCharFormat;
end;

function TRtfListItems.GetNumber(ANumber: Integer): string;
begin
  case FListType of
    ltAlphaUpper: Result := GetAlphaNumber(ANumber, False);
    ltAlphaLower: Result := GetAlphaNumber(ANumber, True);
    ltRomanUpper: Result := GetRomanNumber(ANumber, False);
    ltRomanLower: Result := GetRomanNumber(ANumber, True);
  else
    Result := IntToStr(ANumber);
  end;
end;

function TRtfListItems.GetParagraph: TRtfParagraph;
begin
  raise TRtfException.Create('No Paragraph in list');
end;

function TRtfListItems.GetRomanNumber(ANumber: Integer;
  ALowercase: Boolean): string;
type
  TRomanMapping = record
    Char: string;
    Value: Integer;
  end;
const
  CRomanMapping: array[0..12] of TRomanMapping = (
    (Char: 'M';  Value: 1000),
    (Char: 'CM'; Value: 900),
    (Char: 'D';  Value: 500),
    (Char: 'CD'; Value: 400),
    (Char: 'C';  Value: 100),
    (Char: 'XC'; Value:  90),
    (Char: 'L';  Value:  50),
    (Char: 'XL'; Value:  40),
    (Char: 'X';  Value:  10),
    (Char: 'IX'; Value:   9),
    (Char: 'V';  Value:   5),
    (Char: 'IV'; Value:   4),
    (Char: 'I';  Value:   1)
  );
var
  i, j, LMatches: Integer;
begin
  Result := '';
  for i := Low(CRomanMapping) to High(CRomanMapping) do
  begin
    LMatches := ANumber div CRomanMapping[i].Value;
    for j := 1 to LMatches do
      Result := Result + CRomanMapping[i].Char;
    ANumber := ANumber - LMatches * CRomanMapping[i].Value;
  end;
  if ALowercase then
    Result := lowercase(Result);
end;

function TRtfListItems.GetTotalTextIndent: Double;
begin
  Result := 0;
  if Assigned(FParent) then
    Result := FParent.GetTotalTextIndent;
  Result := Result + FTextindent;
end;

function TRtfListItems.Render: string;
var
  LNumber: Integer;
  i: Integer;
  LItems: TRtfListItems;

  fmt: TRtfCharFormat;

  LListCharacter: string;
  LListCharDefinition: string;
  LTextIndent: Double;
  LFont: TRtfFontDescriptor;
  LFontSize: Double;
var
  LStringBuilder: TStringBuilder;
begin
  LStringBuilder := TStringBuilder.Create;
  try
    LNumber := FStartNumber;
    for i := 0 to FItems.Count -1 do
    begin
      LItems := nil;
      if (FItems[i] is TRtfListItems) then
      begin
        LItems := FItems[i] as TRtfListItems;
      end;
      if Assigned(LItems) then
      begin
        if (LItems.ListType in [ltAlphaUpper, ltAlphaLower, ltRomanUpper, ltRomanLower, ltArabicNum]) then
        begin
          LItems.NumberPrefix := Self.NumberPrefix + Self.GetNumber(LNumber) + Self.NumberSeparator;
          LItems.NumberSuffix := Self.NumberSuffix;
        end
      end
      else
      begin
        fmt := Self.GetListCharFormat;

        if fmt.Font >= 0 then
          LFont := fmt.Font
        else
        begin
          if Assigned(FDocument.FDefaultCharFormat) and (FDocument.FDefaultCharFormat.Font >= 0) then
            LFont := FDocument.FDefaultCharFormat.Font
          else
            LFont := FDocument.DefaultFont;
        end;

        if fmt.FontSize > 0 then
          LFontSize := fmt.FontSize
        else
        begin
          if Assigned(FDocument.FDefaultCharFormat) and (FDocument.FDefaultCharFormat.FFontSize > 0) then
            LFontSize := FDocument.FDefaultCharFormat.FFontSize
          else
            LFontSize := DefaultValue.FontSize;
        end;

        LListCharacter := Self.GetListCharacter(LNumber);
        LListCharDefinition := Format('{\*\pn\pnlvlblt\pnf%d\pnfs%d', [LFont, TRtfUtility.pt2HalfPt(LFontSize)]);
        if fmt.FFGColor <> -1 then
          LListCharDefinition := LListCharDefinition + Format('\pncf%d', [fmt.FFGColor]);
        LListCharDefinition := LListCharDefinition + Format('\pnindent0{\pntxtb %s}}', [LListCharacter]);
        LTextIndent := GetTotalTextIndent;

        LStringBuilder.AppendFormat('\nowidctlpar\fi-%d\li%d', [TRtfUtility.pt2Twip(ListIndent), TRtfUtility.pt2Twip(LTextIndent)]);
        LStringBuilder.AppendLine;
        LStringBuilder.Append(LListCharDefinition);

        Inc(LNumber);
      end;

      LStringBuilder.Append(FItems[i].Render);

      if not (FItems[i] is TRtfListItems) then
      begin
        LStringBuilder.AppendLine('\par}{\pard');
      end;
    end;
    Result := LStringBuilder.ToString;
  finally
    LStringBuilder.Free;
  end;
end;

procedure TRtfListItems.SetListType(const Value: TRtfListType);
var
  LFontName: string;
begin
  if FListType in [ltArrow, ltCircle, ltSquare, ltDiamond, ltBullet] then
  begin
    FreeAndNil(FDefaultCharFormat);
  end;

  FListType := Value;

  case FListType of
    ltArrow: LFontName := 'Wingdings';
    ltCircle: LFontName := 'Courier New';
    ltSquare: LFontName := 'Wingdings';
    ltDiamond: LFontName := 'Wingdings';
    ltBullet: LFontName := 'Symbol';
  else
    LFontName := '';
  end;

  if LFontName <> '' then
  begin
    FDefaultCharFormat := TRtfCharFormat.Create(-1, -1, 0);
    FDefaultCharFormat.Font := FDocument.CreateFont(LFontName);
    FDefaultCharFormat.FontSize := 10;
  end;
end;

{ TRtfList }

constructor TRtfList.Create(AType: TRtfListType; ADoc: TRtfDocument);
begin
  inherited Create(nil, AType, ADoc);
end;

function TRtfList.Render: string;
begin
  Result := '{\pard' + inherited Render + '\par}';
end;

initialization
  DefaultValue.FontSize := 12;
  DefaultValue.Font := 'Times New Roman';
  DefaultValue.MarginLarge := 50;
  DefaultValue.MarginSmall := 50;
end.
