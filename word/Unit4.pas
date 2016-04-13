unit Unit4;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ComCtrls, Vcl.StdCtrls,
  Vcl.Samples.Spin, office_tlb, word_tlb;

type
  TForm4 = class(TForm)
    GroupBox1: TGroupBox;
    GroupBox2: TGroupBox;
    GroupBox3: TGroupBox;
    SpinEdit1: TSpinEdit;
    Label1: TLabel;
    Label2: TLabel;
    Edit1: TEdit;
    Label3: TLabel;
    Label4: TLabel;
    DateTimePicker1: TDateTimePicker;
    Edit2: TEdit;
    Label5: TLabel;
    Edit3: TEdit;
    Label6: TLabel;
    Edit4: TEdit;
    Label7: TLabel;
    Edit5: TEdit;
    Label8: TLabel;
    Edit6: TEdit;
    Edit7: TEdit;
    Label9: TLabel;
    StatusBar1: TStatusBar;
    Button1: TButton;
    procedure Button1Click(Sender: TObject);

  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form4: TForm4;
  WApp: WordApplication;
  Doc: WordDocument;

implementation

{$R *.dfm}

procedure TForm4.Button1Click(Sender: TObject);
var
  docs: Documents;
  pars: Paragraphs;
  par: Paragraph;
  i:integer;
begin
  WApp := CoWordApplication.Create;
  WApp.Visible := true;
  docs := WApp.Documents;
  Doc := docs.Add('Normal', False, EmptyParam, true);

  //Doc.Paragraphs.Item(1).Format.LeftIndent := WApp.CentimetersToPoints(0);
  //Doc.Paragraphs.Item(1).Format.SpaceAfter := 12;
  Doc.Paragraphs.Item(1).Range.Font.Name := 'Liberation Serif';
  Doc.Paragraphs.Item(1).Range.Font.Size := 12;

  //WApp.Selection.ParagraphFormat.LineSpacing := WApp.LinesToPoints(0);
  Doc.Paragraphs.Item(1).Range.Text := '����� �25' +
  #13 + '������������ ��������������' +
  #13 + '������������� ������' +
  #13 + '�� 31 ������� 1998 �. �1274' +
  #13 +
  #13 + '������� � �������� �' + IntToStr(SpinEdit1.Value) +
  #13 +
  #13 + Edit1.Text +
  #13 + //'�������, ���, ��������'+
  #13 + '����' +
  #13 + '��������       ' + DateToStr(DateTimePicker1.Date) +
  #13 + '�����' +
  #13 + '��������       ' + Edit2.Text +
  #13 + '�������� � ���������:' +
  #13 + '����   ' + Edit3.Text +
  #13 + Edit4.Text +
  #13 + '����   ' + Edit5.Text +
  #13 + '���������� ������ ���� �            ��' +
  #13 + '��������  �                       '+Edit6.Text+'         '+DateToStr(date)+
  #13 + '����� ���������������'+
  #13 + '�����������           ' + Edit7.Text +
  #13 +
  #13 +
  #13 + '�������� �� ���� ������� ������� � ������ ���� � �������� �� ��������� ��������� ������'+
  #13 +
  #13 +
  #13 + '���� ������ ' + DateToStr(date)+
  #13 +
  #13 + '�. �.'+
  #13 + '������������ ������'+
  #13 + '������ ����� ������������'+
  #13 + '���������';

  for I := 5 to 18 do
    Doc.Paragraphs.Item(i).Format.Alignment := wdAlignPageNumberLeft;
  for I := 1 to 4 do
    Doc.Paragraphs.Item(i).Format.Alignment := wdAlignPageNumberRight;

  Doc.Paragraphs.Item(6).Range.Font.Bold := 1;
  Doc.Paragraphs.Item(6).Range.Case_ := wdUpperCase;
  Doc.Paragraphs.Item(8).Format.Alignment := wdAlignPageNumberCenter;
  Doc.Paragraphs.Item(9).Format.Alignment := wdAlignPageNumberCenter;
  Doc.Paragraphs.Item(9).Range.Italic := wdToggle;
  Doc.Paragraphs.Item(27).Format.LeftIndent:=WApp.CentimetersToPoints(2);
  Doc.Paragraphs.Item(30).Format.LeftIndent:=WApp.CentimetersToPoints(2);
  Doc.Paragraphs.Item(31).Format.LeftIndent:=WApp.CentimetersToPoints(2);
  Doc.Paragraphs.Item(32).Format.LeftIndent:=WApp.CentimetersToPoints(2);
  WApp.Selection.WholeStory;
  WApp.ActiveDocument.PageSetup.LeftMargin:=WApp.CentimetersToPoints(2);
  doc.Paragraphs.Item(8).Format.LeftIndent := WApp.CentimetersToPoints(0);
  doc.Paragraphs.Item(8).Format.FirstLineIndent := WApp.CentimetersToPoints(0.5);

  //Doc.Paragraphs.Item(5).Range.Font.

end;
end.
