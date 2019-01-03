unit MainForm;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, cxGraphics, cxLookAndFeels,
  cxLookAndFeelPainters, Vcl.Menus, Vcl.StdCtrls, cxButtons, uDMBase, System.Win.ComObj,
  Vcl.Clipbrd, cxControls, cxContainer, cxEdit, cxTextEdit, cxMemo, System.StrUtils,
  System.RegularExpressions, System.Generics.Collections, System.IOUtils, Data.DB,
  Word2000, Vcl.Imaging.jpeg, Vcl.OleServer, Word2010, ZipForge, RVScroll,
  RichView, RVEdit, RVStyle, RVTable, Buttons, CRVData, CRVFData, RVERVData,
  RVItem, RVFuncs, RVTypes, RVGrHandler, System.ImageList, RVGetTextW;

type
  TfrmMain = class(TForm)
    dlgOpen: TOpenDialog;
    btnOpen: TcxButton;
    btnExportXML: TcxButton;
    btnSaveDocFiles: TcxButton;
    rView: TRichViewEdit;
    RVStyle1: TRVStyle;
    procedure btnOpenClick(Sender: TObject);
    procedure btnSaveDocFilesClick(Sender: TObject);
    procedure btnExportXMLClick(Sender: TObject);
  private
    { Private declarations }
//    function CheckString(ASource: string): Boolean;
    procedure ExtractDatafromDoc(ADocName: string);
    procedure ReSaveRtF;
  public
    { Public declarations }
    procedure ReadWordFile;
    procedure ExtractTextFromWordFile(AText: string; AFileName: string; AQuestionIndex: Integer);
    procedure SaveQti;
  end;

var
  frmMain: TfrmMain;
  rsSaveDirectory: string; // = '.\XML_Import\';
  rsHtmlDirectory: string; // = '.\HTML_Import\';

implementation

uses
  XMLStructure;

var
  XMLImport: TSaveXml;
{$R *.dfm}

{ TfrmMain }

procedure TfrmMain.ReadWordFile;
begin
//  if not dlgOpen.Execute then
//    Exit
//  else
  SaveQti;
end;

procedure TfrmMain.ReSaveRtF;

  procedure Save(AFile: string);
  var
    WApp: Variant;
    Doc: Variant;
  begin
    WApp := CreateOleObject('Word.Application');
    WApp.Visible := False;
    Doc := WApp.Documents.Open(AFile);
    if Doc.ReadOnly then
      Exit
    else
    begin
      AFile := StringReplace(AFile, '.doc', '.rtf', [rfReplaceAll]);
      Doc.Range.Select;
      try
        Doc.SaveAs(AFile, wdFormatRTF, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
        AFile := StringReplace(AFile, '.rtf', '.doc', [rfReplaceAll]);
        WApp.Quit;
        DeleteFile(AFile);
      except
        Exit;
      end;
    end;
  end;

var
  Directory, DocList: string;
begin
  Directory := ExtractFilePath(ParamStr(0)) + 'Exam_docs\';
  XMLImport := TSaveXml.Create;
  try
    for DocList in XMLImport.CountFile('Exam_docs\', '*.doc') do
      Save(Directory + DocList);
  finally
    XMLImport.Free;
  end;
end;

procedure TfrmMain.ExtractDatafromDoc(ADocName: string);

  function DeleteTrashSymbols(AString: string): Boolean;
  const
    SQuestion = '^[\d]{1,}\.';
  begin
    if TRegEx.IsMatch(AString, SQuestion) then
      Result := True
    else
      Result := False;
  end;

  function CheckNextSymbol(AIndex: Integer): Boolean;
  begin
    if AIndex = rView.ItemCount - 1 then
      Result := True
    else
    if TRegEx.IsMatch(rView.GetItemTextA(AIndex + 1), '[0-9]{1,1}\.') then
      Result := True;
  end;
const
  Symbols = 15;
var
  I: integer;
  A, ImgName, Folder: string;
  QuestQty: Integer;
  IsQuestion: Boolean;
begin
  if rView.LoadRTF(ADocName) then
  begin
    rView.Format;
    try
      QuestQty := 1;
      Folder := StringReplace(ADocName, '.rtf', '', [rfReplaceAll]);
      for I := 0 to rView.ItemCount - 1 do
      begin
        if DeleteTrashSymbols(rView.GetItemTextW(I)) or IsQuestion then
        begin
          begin
            IsQuestion := True;
            if TRegEx.IsMatch(rView.GetItemTextA(I), '[A-Z]{1,1}\){1,1}') then
              A := A + '#13#10' + rView.GetItemTextA(I)
            else
              A := A + rView.GetItemTextA(I);
          end;
          if rView.GetItemStyle(I) = rvsPicture then
          begin
            ImgName := StringReplace(ExtractFileName(ADocName), '.rtf', '', [rfReplaceAll]);
            ForceDirectories(rsSaveDirectory + ImgName);
            TRVGraphicItemInfo(rView.GetItem(I)).Image.SaveToFile(rsSaveDirectory + '\' + ImgName + '\' + ImgName + '_' + IntToStr(QuestQty) + '.jpeg');
          end;
          if (Length(A) > Symbols) and CheckNextSymbol(I) then
          begin
            ExtractTextFromWordFile(A, ADocName, QuestQty);
            Inc(QuestQty, 1);
            A := '';
            IsQuestion := False;
          end;
        end;
      end;
      DeleteFile(ADocName);
    finally
      rView.Clear;
      rView.Format;
    end;
  end;
end;

procedure TfrmMain.SaveQti;
var
  Directory, DocList: string;
begin
  Directory := ExtractFilePath(ParamStr(0)) + 'Exam_docs_1\';
  for DocList in XMLImport.CountFile('Exam_docs_1\', '*.rtf') do
    ExtractDatafromDoc(Directory + DocList);
end;

procedure TfrmMain.btnExportXMLClick(Sender: TObject);
begin
  ReSaveRtF;
end;

procedure TfrmMain.btnOpenClick(Sender: TObject);
begin
  ReadWordFile;
end;

procedure TfrmMain.btnSaveDocFilesClick(Sender: TObject);

  function RemoveSpecialChars(const str: string): string;
  const
    InvalidChars: set of char = [',', '.', '/', '!', '@', '#', '$', '%', '^', '&', '*', '''', '"', ';', '_', '(', ')', ':', '|', '[', ']', '<', '>', '?'];
  var
    i, Count: Integer;
  begin
    SetLength(Result, Length(str));
    Count := 0;
    for i := 1 to Length(str) do
      if not (str[i] in InvalidChars) then
      begin
        inc(Count);
        Result[Count] := str[i];
      end;
    SetLength(Result, Count);
  end;

var
  FileFolder, FName: string;
begin
  ForceDirectories(ExtractFilePath(ParamStr(0)) + 'Exam_docs\');
  FileFolder := ExtractFilePath(ParamStr(0)) + '\' + 'Exam_docs';
  with DMbase do
  begin
    Exams.Close;
    Exams.Open;
    while not Exams.Eof do
    begin
      if VarIsNull(Exams.FieldByName('_Document_').Value) then
        Exams.Next
      else
      begin
        FName := RemoveSpecialChars(Exams.FieldByName('TmsExam_ID').AsString);
        (Exams.FieldByName('_Document_') as TBlobField).SaveToFile(FileFolder + '\' + FName + '.doc');
        Exams.Next;
      end;
    end;
  end;
end;

procedure TfrmMain.ExtractTextFromWordFile(AText: string; AFileName: string; AQuestionIndex: Integer);

  function ListAnswers(AData: TArray<string>): TList<string>;
  var
    Answers: TList<string>;
    I: Integer;
  begin
    Answers := TList<string>.Create;
    for I := 0 to Length(AData) do
      Answers.Add(AData[I]);
    Result := Answers;
  end;

var
  Ans: TArray<string>;
begin
  XMLImport := TSaveXml.Create;
  try
    Ans := AText.Split(['#13#10']);
    AFileName := StringReplace(ExtractFileName(AFileName), '.rtf', '', [rfReplaceAll]);
    XMLImport.Create.SaveToFile(ListAnswers(Ans), AFileName, AQuestionIndex);
    XMLImport.SaveManifest(AFileName);
  finally
    XMLImport.Free;
  end;
end;

initialization
  ForceDirectories(ExtractFilePath(ParamStr(0)) + 'HTML_Import\');
  rsSaveDirectory := ExtractFilePath(ParamStr(0)) + 'XML_Import\';
  rsHtmlDirectory := ExtractFilePath(ParamStr(0)) + 'HTML_Import\';

end.

