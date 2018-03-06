unit Unit1;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, Buttons,
  StdCtrls, ExcelSheets, BufDataset, db;

type

  { TForm1 }

  TForm1 = class(TForm)
    BitBtn1: TBitBtn;
    BitBtn10: TBitBtn;
    BitBtn11: TBitBtn;
    BitBtn12: TBitBtn;
    BitBtn13: TBitBtn;
    BitBtn14: TBitBtn;
    BitBtn15: TBitBtn;
    BitBtn16: TBitBtn;
    BitBtn17: TBitBtn;
    BitBtn19: TBitBtn;
    BitBtn2: TBitBtn;
    BitBtn20: TBitBtn;
    BitBtn21: TBitBtn;
    BitBtn22: TBitBtn;
    BitBtn23: TBitBtn;
    BitBtn24: TBitBtn;
    BitBtn25: TBitBtn;
    BitBtn3: TBitBtn;
    BitBtn4: TBitBtn;
    BitBtn5: TBitBtn;
    BitBtn6: TBitBtn;
    BitBtn7: TBitBtn;
    BitBtn8: TBitBtn;
    BitBtn9: TBitBtn;
    Button1: TButton;
    Label1: TLabel;
    Sheets1: TSheets;
    procedure BitBtn10Click(Sender: TObject);
    procedure BitBtn11Click(Sender: TObject);
    procedure BitBtn12Click(Sender: TObject);
    procedure BitBtn13Click(Sender: TObject);
    procedure BitBtn14Click(Sender: TObject);
    procedure BitBtn15Click(Sender: TObject);
    procedure BitBtn16Click(Sender: TObject);
    procedure BitBtn17Click(Sender: TObject);
    procedure BitBtn19Click(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure BitBtn20Click(Sender: TObject);
    procedure BitBtn21Click(Sender: TObject);
    procedure BitBtn22Click(Sender: TObject);
    procedure BitBtn23Click(Sender: TObject);
    procedure BitBtn24Click(Sender: TObject);
    procedure BitBtn25Click(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure BitBtn3Click(Sender: TObject);
    procedure BitBtn4Click(Sender: TObject);
    procedure BitBtn5Click(Sender: TObject);
    procedure BitBtn6Click(Sender: TObject);
    procedure BitBtn7Click(Sender: TObject);
    procedure BitBtn8Click(Sender: TObject);
    procedure BitBtn9Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
  private

  public

  end;

var
  Form1: TForm1;

implementation

{$R *.lfm}

{ TForm1 }

procedure TForm1.BitBtn1Click(Sender: TObject);
begin
  Sheets1.VisualizarExcel;
end;





procedure TForm1.BitBtn2Click(Sender: TObject);
begin
  Sheets1.OcutarExcel;
end;

procedure TForm1.BitBtn3Click(Sender: TObject);
begin
  Sheets1.VisualizarExcel;
  Sheets1.CelulaValor('CÓDIGO'); //imprime na A1
end;

procedure TForm1.BitBtn4Click(Sender: TObject);
begin
  Sheets1.IrParaCelula('B1');   //vai para b1
  Sheets1.CelulaValor('CLIENTE');
end;

procedure TForm1.BitBtn5Click(Sender: TObject);
begin
  Sheets1.CelulaFonte('Arial'); //Muda b1 para arial
  Sheets1.IrParaCelula('A1');   //vai para A1
  Sheets1.CelulaFonte('Arial'); //Muda b1 para arial

end;

procedure TForm1.BitBtn6Click(Sender: TObject);
begin
  Sheets1.CelulaNegrito(TRUE);  //POE NEGRITO NA CELULA ATUAL
  Sheets1.IrParaCelula('B1');   //vai para B1
  Sheets1.CelulaNegrito(TRUE);  //NEGRITO CELULA B1


end;

procedure TForm1.BitBtn7Click(Sender: TObject);
begin
  Sheets1.CelulaSize(14); //Aumenta fonte celula B1
  Sheets1.IrParaCelula('A1');  //vai para A1
  Sheets1.CelulaSize(14); //Aumenta Fonte
end;

procedure TForm1.BitBtn8Click(Sender: TObject);
begin
  Sheets1.CelulaColor(ClRed);
  Sheets1.IrParaCelula('B1');
  Sheets1.CelulaColor(ClRed);

end;

procedure TForm1.BitBtn9Click(Sender: TObject);
begin
  Sheets1.CelulaBackGround(ClYELLOW);
  Sheets1.IrParaCelula('A1');
  Sheets1.CelulaBackGround(ClYELLOW);
end;

procedure TForm1.Button1Click(Sender: TObject);
begin
  Sheets1.Abrir('D:\TESTE82\13salario2017.xlsx');
end;

procedure TForm1.BitBtn19Click(Sender: TObject);
begin
  Sheets1.CelulaBordas(TRUE,TRUE,TRUE,TRUE);
  Sheets1.IrParaCelula('B1');
  Sheets1.CelulaBordas(TRUE,TRUE,TRUE,TRUE);
  Sheets1.ColunaAutoAjuste;
end;

procedure TForm1.BitBtn10Click(Sender: TObject);
begin
 Sheets1.Cells(2,1,10);
 Sheets1.Cells(2,2,'Humberto');
 Sheets1.Cells(3,1,10);
 Sheets1.Cells(3,2,'Sales');

 //criar coluna Valor;
 Sheets1.Cells(1,3,'Valor');
 Sheets1.Cells(2,3,16);
 Sheets1.Cells(3,3,17);
 Sheets1.ColunaSomar('C2:C3');

 Sheets1.IrParaCelula('C4');
 Sheets1.CelulaNegrito(True);
 Sheets1.CelulaMascara('R$#.##0,00');
 Sheets1.IrParaCelula('C1');
 Sheets1.CelulaColor(clred);
 Sheets1.CelulaBackGround(clblack);

end;

procedure TForm1.BitBtn11Click(Sender: TObject);
begin
  Sheets1.FaixaBackGround('A1:C1',CLYELLOW);
end;

procedure TForm1.BitBtn12Click(Sender: TObject);
begin
    Sheets1.FaixaMesclar('D1:D4');
end;

procedure TForm1.BitBtn13Click(Sender: TObject);
begin
  Sheets1.FaixaValor('H1:G1','FAIXA VALOR');
  Sheets1.FaixaBordas('H1:G1',TRUE,TRUE,TRUE,TRUE)
end;

procedure TForm1.BitBtn14Click(Sender: TObject);
var
  buf : TBufDataSet;
begin
  buf :=  TBufDataset.Create(Nil);
  With buf do
        begin
            FieldDefs.Add('codigo',ftInteger);
            FieldDefs.Add('Nome',ftString,50);
            FieldDefs.Add('Valor',ftFloat);
            CreateDataset;
            Insert;
            Fields.Fieldbyname('Codigo').AsInteger := 1;
            Fields.Fieldbyname('Nome').AsString    := 'Humberto';
            Fields.Fieldbyname('Valor').AsFloat    := 17;
            Post;
            Insert;
            Fields.Fieldbyname('Codigo').AsInteger := 2;
            Fields.Fieldbyname('Nome').AsString    := 'Denis';
            Fields.Fieldbyname('Valor').AsFloat    := 21;
            Post;
            Insert;
            Fields.Fieldbyname('Codigo').AsInteger := 3;
            Fields.Fieldbyname('Nome').AsString    := 'Patricia';
            Fields.Fieldbyname('Valor').AsFloat    := 41;
            Post;
            Sheets1.Dataset(Buf,5,2);
            Sheets1.ColunaSomar('D6:D8');
            FREE;
        end;
end;

procedure TForm1.BitBtn15Click(Sender: TObject);
begin
  Sheets1.visualizarImpressao;
end;

procedure TForm1.BitBtn16Click(Sender: TObject);
begin

  Sheets1.Salvar('d:\planilha.xlsx');
end;

procedure TForm1.BitBtn17Click(Sender: TObject);
begin
  Sheets1.imprimir;
end;


procedure TForm1.BitBtn20Click(Sender: TObject);
begin
  Sheets1.FaixaColunaAutoAjuste('A1:B1');
end;

procedure TForm1.BitBtn21Click(Sender: TObject);
begin
  Sheets1.ColunaWidth(40);
  sheets1.LinhaHeight(40);
end;

procedure TForm1.BitBtn22Click(Sender: TObject);
begin
   Sheets1.FaixaLinhaAutoAjuste('A1:B1');
end;

procedure TForm1.BitBtn23Click(Sender: TObject);
begin
  Sheets1.IrParaCelula('A1');
  SHOWMESSAGE(Sheets1.GetCelulaValor);
end;

procedure TForm1.BitBtn24Click(Sender: TObject);
begin
  Sheets1.ColunaAutoAjuste;
end;

procedure TForm1.BitBtn25Click(Sender: TObject);
begin
    Sheets1.CelulaAlignHorizontal(AHCenter);
    Sheets1.CelulaAlignVertical(AVCenter);
end;

end.

