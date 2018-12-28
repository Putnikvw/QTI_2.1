unit XMLStructure;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, cxGraphics, cxLookAndFeels,
  cxLookAndFeelPainters, Vcl.Menus, Vcl.StdCtrls, cxButtons, uDMBase, System.Win.ComObj,
  Vcl.Clipbrd, cxControls, cxContainer, cxEdit, cxTextEdit, cxMemo, System.StrUtils,
  System.RegularExpressions, System.Generics.Collections, Xml.Xmldom, Xml.XMLIntf,
  Xml.XMLDoc, System.IOUtils, System.DateUtils, System.ZLib, MainForm;

type
  TSaveXml = class(TObject)
  private

  public
    function CountFile(ADirectory: string; const AFileType: string = '*_*.xml'): TList<string>;
    procedure SaveToFile(AList: TList<string>; AFileName: string; ALevel: integer);
    procedure SaveManifest(AFolder: string);
    procedure SaveData(AGuid: TDictionary<string, string>; ADirectory: string);
  end;

implementation

//var
//  rsSaveDirectory: string; // = '.\XML_Import\';
{ TSaveXml }

function TSaveXml.CountFile(ADirectory: string; const AFileType: string = '*_*.xml'): TList<string>;
var
  sr: TSearchRec;
begin
  Result := TList<string>.Create;
  ChDir(ExtractFilePath(ParamStr(0)) + '\' + ADirectory);
  if FindFirst(AFileType, faAnyFile, sr) = 0 then
    repeat
      Result.Add(sr.Name);
    until FindNext(sr) <> 0;
  FindClose(sr);
end;

procedure TSaveXml.SaveData(AGuid: TDictionary<string, string>; ADirectory: string);
const
  NameArr: array[0..2] of string = ('A-Parm', 'B-Parm', 'C-Parm');
var
  XML: IXMLDocument;
  Root, Child: IXMLNode;
  GUID: string;
  I: integer;
begin
  XML := TXMLDocument.Create(Application) as IXMLDocument;
  try
    XML.Active := True;
    XML.Encoding := 'utf-8';
    XML.Version := '1.0';
    Root := XML.AddChild('usageData', 'http://www.imsglobal.org/xsd/imsqti_v2p1');
    Root.Attributes['xmlns:xsi'] := 'http://www.w3.org/2001/XMLSchema-instance';
    Root.Attributes['glossary'] := 'http://www.imsglobal.org/question/qti_2p0/glossaries/item_statistics';
    Root.Attributes['xsi:schemaLocation'] := 'http://www.imsglobal.org/xsd/imsqti_v2p1 http://www.imsglobal.org/xsd/imsqti_v2p1.xsd';
    for GUID in AGuid.Values do
    begin
      for I := 0 to Length(NameArr) - 1 do
      begin
        Child := Root.AddChild('ordinaryStatistic');
        Child.Attributes['name'] := NameArr[I];
        Child.Attributes['context'] := 'http://ntweb.ets.org/itemStats/Context/XVZ2004_J1B';
        Child.Attributes['caseCount'] := '689325';
        Child.Attributes['stdError'] := '0.0022';
        Child.Attributes['lastUpdated'] := '';

        Child.AddChild('targetObject').Attributes['identifier'] := GUID;
        Child.AddChild('value');
      end;
    end;
    XML.SaveToFile(rsSaveDirectory + '\' + ADirectory + '\' + 'usageData.xml');
  finally
    XML := nil;
  end;

end;

procedure TSaveXml.SaveManifest(AFolder: string);

  function RemoveNumbers(const AString: string): string;
  begin
//    Result := '';
//    for C in AString do
//    begin
//      if not CharInSet(C, ['0'..'9']) then
//      begin
//        Result := Result + C;
//      end;
//    end;
    Result := StringReplace(Result, '.xml', '', [rfReplaceAll]);
  end;

  function ReturnGUID: string;
  begin
    Result := StringReplace(StringReplace(TGUID.NewGuid.ToString, '{', '', []), '}', '', []);
  end;

var
  XML: IXMLDocument;
  Root, Resources, RS, Metadata: IXMLNode;
  FileName, GUID: string;
  GuidList: TDictionary<string, string>;
begin
  XML := TXMLDocument.Create(Application) as IXMLDocument;
  try
    XML.Active := True;
    XML.Encoding := 'utf-8';
    XML.Version := '1.0';
    Root := XML.AddChild('manifest', 'http://www.imsglobal.org/xsd/imscp_v1p1');
    Root.Attributes['xmlns:xsi'] := 'http://www.w3.org/2001/XMLSchema-instance';
    Root.Attributes['xmlns:imsqti'] := 'http://www.imsglobal.org/xsd/imsqti_v2p1';
    Root.Attributes['xmlns:imsmd'] := 'http://www.imsglobal.org/xsd/imsmd_v1p2';
    Root.Attributes['xsi:schemaLocation'] := 'http://www.imsglobal.org/xsd/imscp_v1p1 http://www.imsglobal.org/xsd/imscp_v1p1.xsd ' + 'http://www.imsglobal.org/xsd/imsmd_v1p2 http://www.imsglobal.org/xsd/imsmd_v1p2p2.xsd' + 'http://www.imsglobal.org/xsd/imsqti_v2p1 http://www.imsglobal.org/xsd/imsqti_v2p1.xsd';
    Root.Attributes['identifier'] := '';
    Root.AddChild('metadata');
    Root.AddChild('organizations');
    Resources := Root.AddChild('resources');
    GuidList := TDictionary<string, string>.Create;
    try
      for FileName in CountFile('XML_Import\' + AFolder + '\') do
      begin
        GuidList.AddOrSetValue(FileName, ReturnGUID);
        GuidList.TryGetValue(FileName, GUID);
        RS := Root.ChildNodes['resources'].AddChild('resource');
        RS.Attributes['identifier'] := GUID;
        RS.Attributes['href'] := FileName;
        RS.Attributes['type'] := 'imsqti_item_xmlv2p1';
        Metadata := RS.AddChild('metadata');
        Metadata := Metadata.AddChild('imsmd:lom').AddChild('imsmd:general');
        Metadata.AddChild('imsmd:identifier').NodeValue := GUID;
        Metadata.AddChild('imsmd:title').AddChild('imsmd:langstring').Attributes['xml:lang'] := 'en';
        Metadata.ChildNodes['imsmd:title'].ChildNodes['imsmd:langstring'].NodeValue := FileName;
        Metadata.AddChild('imsmd:description').AddChild('imsmd:langstring').Attributes['xml:lang'] := 'en';
        Metadata.AddChild('imsmd:keywords').AddChild('imsmd:langstring').Attributes['xml:lang'] := 'en';
        Metadata.AddChild('imsmd:author').AddChild('imsmd:langstring').Attributes['xml:lang'] := 'en';
        Metadata := Metadata.ParentNode;
        Metadata := Metadata.AddChild('imsmd:metametadata');
        Metadata.AddChild('imsmd:metadatascheme').NodeValue := 'LOMv1.0';
        Metadata.AddChild('imsmd:metadatascheme').NodeValue := 'QTIv2.1';
        Metadata.AddChild('imsmd:language').NodeValue := 'en';
        Metadata.ParentNode.AddChild('imsmd:technical').AddChild('imsmd:format').NodeValue := 'text/x-imsqti-item-xml';
        Metadata := Metadata.ParentNode.ParentNode;
        Metadata := Metadata.AddChild('imsqti:qtiMetadata');
        Metadata.AddChild('imsqti:timeDependent').NodeValue := 'false';
        Metadata.AddChild('imsqti:feedbackType').NodeValue := 'nonadaptive';
        Metadata.AddChild('imsqti:solutionAvailable').NodeValue := 'true';
        Metadata.AddChild('imsqti:toolName').NodeValue := 'FastTestWeb';
        Metadata.AddChild('imsqti:toolVersion').NodeValue := '3.75.8';
        Metadata.AddChild('imsqti:toolVendor').NodeValue := '4ROI';

        Metadata.ParentNode.ParentNode.AddChild('file').Attributes['href'] := FileName;
      end;
      XML.SaveToFile(rsSaveDirectory + '\' + AFolder + '\' + 'imsmanifest.xml');
      SaveData(GuidList, AFolder);
    finally
      GuidList.Free;
    end;
  finally
    XML := nil;
  end;
end;

procedure TSaveXml.SaveToFile(AList: TList<string>; AFileName: string; ALevel: integer);
var
  XML: IXMLDocument;
  Root: IXMLNode;
  I: integer;
begin
  XML := TXMLDocument.Create(Application) as IXMLDocument;
  try
    XML.Active := True;
    XML.Encoding := 'utf-8';
    XML.Version := '1.0';
    Root := XML.AddChild('assessmentItem', 'http://www.imsglobal.org/xsd/imsqti_v2p1');
    Root.Attributes['xmlns:xsi'] := 'http://www.w3.org/2001/XMLSchema-instance';
    Root.Attributes['xsi:schemaLocation'] := 'http://www.imsglobal.org/xsd/imsqti_v2p1 http://www.imsglobal.org/xsd/imsqti_v2p1.xsd';
    Root.Attributes['adaptive'] := 'false';
    Root.Attributes['timeDependent'] := 'false';
    Root.Attributes['identifier'] := AFileName;
    Root.Attributes['title'] := AList.Items[0];

    Root.AddChild('responseDeclaration').AddChild('correctResponse').AddChild('value').NodeValue := AFileName + '_R_' + IntToStr(ALevel);
    Root.ChildNodes['responseDeclaration'].Attributes['baseType'] := 'identifier';
    Root.ChildNodes['responseDeclaration'].Attributes['identifier'] := 'RESPONSE';
    Root.ChildNodes['responseDeclaration'].Attributes['cardinality'] := 'cardinality';

    Root.ChildNodes['responseDeclaration'].AddChild('mapping').AddChild('mapEntry').Attributes['mappedValue'] := '1.0';
    Root.ChildNodes['responseDeclaration'].ChildNodes['mapping'].Attributes['defaultValue'] := '0';
    Root.ChildNodes['responseDeclaration'].ChildNodes['mapping'].Attributes['upperBound'] := '1.0';
    Root.ChildNodes['responseDeclaration'].ChildNodes['mapping'].Attributes['lowerBound'] := '0';
    Root.ChildNodes['responseDeclaration'].ChildNodes['mapping'].ChildNodes['mapEntry'].Attributes['mapKey'] := AFileName + '_R_' + IntToStr(ALevel);

    Root.ChildNodes['outcomeDeclaration'].AddChild('defaultValue').AddChild('value').NodeValue := '0';
    Root.ChildNodes['outcomeDeclaration'].Attributes['cardinality'] := 'single';
    Root.ChildNodes['outcomeDeclaration'].Attributes['baseType'] := 'integer';
    Root.ChildNodes['outcomeDeclaration'].Attributes['identifier'] := 'SCORE';

    Root.AddChild('stylesheet');
    Root.ChildNodes['stylesheet'].Attributes['type'] := 'text/css';
    Root.ChildNodes['stylesheet'].Attributes['href'] := '';

    Root.AddChild('itemBody').AddChild('p').NodeValue := AList.Items[0];
    Root.ChildNodes['itemBody'].AddChild('choiceInteraction').Attributes['responseIdentifier'] := 'RESPONSE';
    Root.ChildNodes['itemBody'].ChildNodes['choiceInteraction'].Attributes['shuffle'] := '';
    Root.ChildNodes['itemBody'].ChildNodes['choiceInteraction'].Attributes['maxChoices'] := '1';
    for I := 1 to AList.Count - 1 do
      Root.ChildNodes['itemBody'].ChildNodes['choiceInteraction'].AddChild('simpleChoice').AddChild('p').NodeValue := AList.Items[I];
    for I := 0 to Root.ChildNodes['itemBody'].ChildNodes['choiceInteraction'].ChildNodes.Count - 1 do
      Root.ChildNodes['itemBody'].ChildNodes['choiceInteraction'].ChildNodes[I].Attributes['identifier'] := AFileName + '_R_' + IntToStr(I);

    Root.AddChild('responseProcessing').Attributes['template'] := 'http://www.imsglobal.org/question/qti_v2p1/rptemplates/map_response';
    ForceDirectories(rsSaveDirectory + AFileName);
    XML.SaveToFile(rsSaveDirectory + '\' + AFileName + '\' + AFileName + '_' + IntToStr(ALevel) + '.xml');
  finally
    XML := nil;
  end;
end;

end.

