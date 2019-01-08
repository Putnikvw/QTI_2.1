unit MainForm;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, cxGraphics, cxLookAndFeels,
  cxLookAndFeelPainters, Vcl.Menus, Vcl.StdCtrls, cxButtons, uDMBase, System.Win.ComObj,
  Vcl.Clipbrd, cxControls, cxContainer, cxEdit, cxTextEdit, cxMemo, System.StrUtils,
  System.RegularExpressions, System.Generics.Collections, System.IOUtils, Data.DB,
  Word2000, Vcl.Imaging.jpeg, Vcl.OleServer, Word2010, RVScroll, RichView,
  RVEdit, RVStyle, RVTable, Buttons, CRVData, CRVFData, RVERVData, RVItem,
  RVFuncs, RVTypes, RVGrHandler, System.ImageList, RVGetTextW, dxSkinsCore,
  FireDAC.Stan.Intf, FireDAC.Stan.Option, FireDAC.Stan.Error, FireDAC.UI.Intf,
  FireDAC.Phys.Intf, FireDAC.Stan.Def, FireDAC.Stan.Pool, FireDAC.Stan.Async,
  FireDAC.Phys, FireDAC.Phys.SQLite, FireDAC.Phys.SQLiteDef, FireDAC.Stan.ExprFuncs,
  FireDAC.VCLUI.Wait, FireDAC.Stan.Param, FireDAC.DatS, FireDAC.DApt.Intf,
  FireDAC.Comp.DataSet, FireDAC.Comp.Client, ZipForge;

type
  TfrmMain = class(TForm)
    dlgOpen: TOpenDialog;
    btnOpen: TcxButton;
    btnExportXML: TcxButton;
    btnSaveDocFiles: TcxButton;
    rView: TRichViewEdit;
    RVStyle1: TRVStyle;
    btnExtractItems: TButton;
    procedure btnOpenClick(Sender: TObject);
    procedure btnSaveDocFilesClick(Sender: TObject);
    procedure btnExportXMLClick(Sender: TObject);
    procedure btnExtractItemsClick(Sender: TObject);
  private
    { Private declarations }
//    function CheckString(ASource: string): Boolean;
    procedure ExtractDatafromDoc(ADocName: string);
    procedure ReSaveRtF;
    procedure ExportDoc(const DirName, FieldID: string; DataSet: TDataSet);
    procedure AddZip(ArchiveName: string);
    procedure DeleteFilesinDir(APath: string);
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
var
  WApp: Variant;

  procedure Save(AFile: string);
  var
    Doc: Variant;
  begin
    Doc := WApp.Documents.Open(AFile);
    try
      if not Doc.ReadOnly then
      begin
        AFile := StringReplace(AFile, '.doc', '.rtf', [rfReplaceAll]);
        Doc.Range.Select;
        try
          Doc.SaveAs(AFile, wdFormatRTF, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam,
          EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
          AFile := StringReplace(AFile, '.rtf', '.doc', [rfReplaceAll]);
          DeleteFile(AFile);
        except
          Exit;
        end;
      end;
    finally
      Doc.Close(0);
    end;
  end;

var
  Directory, DocList: string;
  list: TList<string>;
begin
  Directory := ExtractFilePath(ParamStr(0)) + 'Items_docs\';
  XMLImport := TSaveXml.Create;
  try
    WApp := CreateOleObject('Word.Application');
    WApp.Visible := False;
    list := XMLImport.CountFile('Items_docs\', '*.doc');
    try
      for DocList in list do
        Save(Directory + DocList);
    finally
      list.Free;
    end;
  finally
    WApp.Quit;
    XMLImport.Free;
  end;
end;

procedure TfrmMain.ExportDoc(const DirName, FieldID: string; DataSet: TDataSet);

  procedure _ReplaceData(Source: PChar; len: Integer);
  const
    ps: Char = '\';
    ps0 = '\\perssigmadb\tmsdata\tmsdatabu\tms5d\exams.dot';
    ps1 = '\\pstb\vol1\software\sigma\tms\tms5data\items.dot';
    ps2 = '\\pstb\vol1\software\sigma\tms\tms5data\exams.dot';
    ps3 = '\\pstb\vol1\software\sigma\tms\tms5data\items.dot';
    pss: array[0..3] of PChar = (ps0, ps1, ps2, ps3);
  var
    P, CheckSource: PChar;
    i, j, fcount: Integer;
  begin
    for i := 0 to len - 1 do
    begin
      if Source^ = ps then
        for j := 0 to Length(pss) - 1 do
        begin
          P := pss[j];
          fcount := Length(pss[j]);
          CheckSource := Source;
          while (CheckSource^ = P^) or (Char(Word(CheckSource^) or $20) = P^) do
          begin
            Inc(CheckSource);
            Inc(P);
            Dec(fcount);
          end;

          if (fcount <= 0) then
          begin
            Source^ := #0;
            exit;
          end;
        end;
      Inc(Source);
    end;
  end;

  procedure _SaveToFile(const FileName: string; Blob: TBlobField);
  var
    fs: TFileStream;
    bs: TStream;
    Buf: RawByteString;
    len, lenP: Integer;
  begin
    bs := Blob.DataSet.CreateBlobStream(Blob, bmRead);
    try
      len := bs.Size;
      SetLength(Buf, len);
      bs.Read(Pointer(Buf)^, len);
    finally
      bs.Free;
    end;
    _ReplaceData(Pointer(Buf), len div 2);

    fs := TFileStream.Create(FileName, fmCreate);
    try
      fs.Write(Pointer(Buf)^, len);
    finally
      fs.Free;
    end;
  end;

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
  FileFolder, FileName: string;
  fldID, fldDoc: TField;
begin
  FileFolder := ExtractFilePath(ParamStr(0)) + DirName + '\';
  ForceDirectories(FileFolder);

  with DataSet do
  begin
    Close;
    Open;
    fldID := FieldByName(FieldID);
    fldDoc := FieldByName('_Document_');
    while not Eof do
    begin
      if VarIsNull(fldDoc.Value) then
        Next
      else
      begin
        FileName := RemoveSpecialChars(fldID.AsString);
        _SaveToFile(FileFolder + FileName + '.doc', TBlobField(fldDoc));
        Next;
      end;
    end;
  end;
end;

procedure TfrmMain.ExtractDatafromDoc(ADocName: string);

  function CheckDocument(AString: string): Boolean;
  const
    SQuestion = '^[\d]{1,}\.';
    Items = '^Item ID:[0-9]{5,5}';
    Domain = 'Domain ID:';
    ItTable = '[A-Z]{1,1}[.]*';
    StrLenght = 25;
  begin
    if TRegEx.IsMatch(AString, SQuestion) or TRegEx.IsMatch(AString, Items) or TRegEx.IsMatch(AString, Domain) or
      (TRegEx.IsMatch(AString, ItTable) and (Length(AString) >= StrLenght)) then
      Result := True
    else
      Result := False;
  end;

  function DeleteSymbols(AText: string): string;
  begin
    AText := TRegEx.Replace(AText, '^\s{1,}', '', [roIgnoreCase]);
    AText := TRegEx.Replace(AText, '^[A-K]{1,1}\){1,1}', '', [roIgnoreCase]);
    AText := TRegEx.Replace(AText, '^\s{1,}', '', [roIgnoreCase]);
    AText := TRegEx.Replace(AText, '^[0-9]{1,1}\.', '', [roIgnoreCase]);
    Result := AText;
  end;

  function CheckNextSymbol(AIndex: Integer): Boolean;
  begin
    if AIndex = rView.ItemCount - 1 then
      Result := True
    else if TRegEx.IsMatch(rView.GetItemTextA(AIndex + 1), '[0-9]{1,1}\.') then
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
        if CheckDocument(rView.GetItemTextW(I)) or IsQuestion then
        begin
          begin
            IsQuestion := True;
            a := Trim(A);
            A := TRegEx.Replace(A, '^[0-9]{1,1}\.', '', [roIgnoreCase]);
            if TRegEx.IsMatch(rView.GetItemTextA(I), '[A-Z]{1,1}\){1,1}') then
              A := A + '#13#10' + DeleteSymbols(rView.GetItemTextA(I))
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
            A := TRegEx.Replace(A, '\s{2,}', '', [roIgnoreCase]);
            A := TRegEx.Replace(A, '^[0-9]{1,}\.', '', [roIgnoreCase]);
            ExtractTextFromWordFile(A, ADocName, QuestQty);
            Inc(QuestQty, 1);
            A := '';
            IsQuestion := False;
          end;
        end;
      end;
      AddZip(ADocName);
//      DeleteFile(ADocName);
    finally
      rView.Clear;
      rView.Format;
    end;
  end;
end;

procedure TfrmMain.SaveQti;
const
  Exam_1 = 'Exam_docs_1\';
  Exam = 'Exam_docs\';
  Items = 'Items_docs\';
var
  Directory, DocList: string;
  FileList: TList<string>;
begin
  Directory := ExtractFilePath(ParamStr(0)) + Items;
  FileList := XMLImport.CountFile(Items, '*.rtf');
  try
    for DocList in FileList do
      ExtractDatafromDoc(Directory + DocList);
    MessageBox(0, PChar('OK!'), PChar(''), MB_ICONINFORMATION or MB_OK);
  finally
    FileList.Free;
  end;
end;

procedure TfrmMain.AddZip(ArchiveName: string);
begin
  ArchiveName := StringReplace(ExtractFileName(ArchiveName), '.rtf', '', [rfReplaceAll]);
  with TZipForge.Create(nil) do
  begin
    try
      FileName := ArchiveName + '.zip';
      OpenArchive(fmCreate);
      AddFiles(rsSaveDirectory + ArchiveName + '\' + '*.*');
      CloseArchive();
      DeleteFilesinDir(rsSaveDirectory + ArchiveName + '\');
    finally
      Free;
    end;
  end;
end;

procedure TfrmMain.btnExportXMLClick(Sender: TObject);
begin
  ReSaveRtF;
end;

procedure TfrmMain.btnExtractItemsClick(Sender: TObject);
begin
  ExportDoc('Items_docs', 'Item_ID', DMbase.Items);
  MessageBox(0, PChar('OK!'), PChar(''), MB_ICONINFORMATION or MB_OK);

end;

procedure TfrmMain.btnOpenClick(Sender: TObject);
begin
  ReadWordFile;
end;

procedure TfrmMain.btnSaveDocFilesClick(Sender: TObject);
begin
  ExportDoc('Exam_docs', 'TmsExam_ID', DMbase.Exams);
  MessageBox(0, PChar('OK!'), PChar(''), MB_ICONINFORMATION or MB_OK);
end;

procedure TfrmMain.DeleteFilesinDir(APath: string);
var
  sr: TSearchRec;
  FilesList: TArray<string>;
  I: Integer;
begin
  FilesList := ['*.xml', '*.jpeg'];
  for I := 0 to Length(FilesList) - 1 do
    if FindFirst(APath + FilesList[I], faAnyFile, sr) = 0 then
      repeat
        DeleteFile(sr.Name);
      until FindNext(sr) <> 0;
  FindClose(sr);
end;

procedure TfrmMain.ExtractTextFromWordFile(AText: string; AFileName: string; AQuestionIndex: Integer);

  procedure ListAnswers(AData: TArray<string>; var AList: TList<string>);
  var
    I: Integer;
  begin
    for I := 0 to Length(AData) - 1 do
      AList.Add(AData[I]);
  end;

var
  Ans: TArray<string>;
  List: TList<string>;
begin
  XMLImport := TSaveXml.Create;
  try
    Ans := AText.Split(['#13#10']);
    AFileName := StringReplace(ExtractFileName(AFileName), '.rtf', '', [rfReplaceAll]);
    List := TList<string>.Create;
    try
      ListAnswers(Ans, List);
      XMLImport.SaveToFile(List, AFileName, AQuestionIndex);
      XMLImport.SaveManifest(AFileName);
    finally
      List.Free;
    end;
  finally
    XMLImport.Free;
  end;
end;

initialization
  ForceDirectories(ExtractFilePath(ParamStr(0)) + 'HTML_Import\');
  rsSaveDirectory := ExtractFilePath(ParamStr(0)) + 'XML_Import\';
  rsHtmlDirectory := ExtractFilePath(ParamStr(0)) + 'HTML_Import\';

end.

