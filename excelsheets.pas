unit ExcelSheets;

{$mode objfpc}{$H+}

interface

uses
  Classes, {$IFDEF MSWINDOWS}comobj, {$ENDIF}variants, db, SysUtils, Graphics, LResources, Dialogs;


Type
TalignVer = (AVTop,AVCenter,AVBottom);
TAlignHor = (AHCenter,AHRight,AHLeft);

{ TStylo }

TStylo = Class
private
    FAlignHor: TAlignHor;
    FAlignVer: TalignVer;
    FAlturaLinha: Integer;
    FbackGround: TColor;
    FColor: TColor;
    FFonte: String;
    FNegrito: Boolean;
    FSize: Byte;
    procedure SetAlignHor(AValue: TAlignHor);
    procedure SetAlignVer(AValue: TalignVer);
    procedure SetAlturaLinha(AValue: Integer);
    procedure SetbackGround(AValue: TColor);
    procedure SetColor(AValue: TColor);
    procedure SetFonte(AValue: String);
    procedure SetNegrito(AValue: Boolean);
    procedure SetSize(AValue: Byte);
published
    property Fonte : String read FFonte write SetFonte ;
    property Size  : Byte read FSize write SetSize default 12;
    property AlignHor : TAlignHor read FAlignHor write SetAlignHor default AHLeft;
    property AlignVer : TalignVer read FAlignVer write SetAlignVer default AVCenter;
    property Negrito : Boolean read FNegrito write SetNegrito default false;
    property backGround : TColor read FbackGround write SetbackGround default cldefault;
    property Color : TColor read FColor write SetColor default cldefault;
    property AlturaLinha : Integer read FAlturaLinha write SetAlturaLinha default 23;
end;
TFaixa = Array[1..2] of Variant;

{ TSheets }

TSheets = Class(TComponent)

  private
    FCelula: Variant;
    ObjExcel,Sheet : Variant;
    FNome : Variant;
    FSheetName : Variant;
    FVisualizaExcel : boolean; //faz parte de visualizar/ocutarexcel

    function GetSheetName: Variant;
    function GetNome: Variant;
    procedure SetCelula(AValue: Variant);
    procedure SetNome(AValue: Variant);
    procedure SetSheetName(AValue: Variant);
    function Alfabetico(qtd: integer): String;
  public
    function GetCelulaValor: Variant;
    function GetColunaWidth: Variant;
    function GetLinhaHeight: Variant;
    procedure CelulaAlignHorizontal(AValue: TAlignHor);
    procedure CelulaAlignVertical(AValue: TalignVer);
    procedure CelulaBackGround(AValue: TColor);
    procedure CelulaColor(AValue: TColor);
    procedure CelulaFonte(AValue: Variant);
    procedure CelulaMascara(AValue: Variant);
    procedure CelulaNegrito(AValue: Boolean);
    procedure CelulaSize(AValue: Byte);
    procedure CelulaValor(AValue: Variant);
    procedure CelulaReplace(AValueAtual,AValueNovo : Variant);
    procedure CelulaBordas(direita, esquerda, superior, inferior: Boolean);
    procedure CelulaSomar(Item1,Item2: Variant);
    procedure CelulaFormatoCurrency(Item1, Item2: Variant);
    procedure CelulaFormatoTexto(Item1, Item2: Variant);
    procedure ColunaWidth(AValue: Variant);
    Function ColunaUltima : Integer;
    procedure ColunaAutoAjuste;
    procedure ColunaInserir(NumCol: Integer);
    procedure ColunaApagar(NumCol: Integer);
    procedure ColunaSomar(Faixa: Variant);
    procedure LinhaSomar(Faixa:Variant);
    procedure LinhaHeight(AValue: Variant);
    function LinhaUltima : Integer;
    procedure LinhaAutoAjuste;
    procedure LinhaApagar(NumLin: Integer);
    procedure LinhaInserir(NumLin: Integer);
    procedure FaixaFonte(Faixa: Variant; FontName: Variant);
    procedure FaixaSize(Faixa: Variant; FontSize: byte);
    procedure FaixaNegrito(Faixa: Variant; Negrito: Boolean);
    procedure FaixaColor(Faixa: Variant; Cor: TColor);
    procedure FaixaBackGround(Faixa: Variant; Cor: TColor);
    procedure FaixaMascara(Faixa: Variant; Mascara: Variant);
    procedure FaixaColunaAutoAjuste(Faixa: Variant);
    procedure FaixaLinhaAutoAjuste(Faixa: Variant);
    procedure FaixaLinhaHeight(Faixa: Variant; height: Variant);
    procedure FaixaColunaWidth(Faixa: Variant; Width: Variant);
    procedure FaixaAlignHorizontal(Faixa: Variant; Alinhamento: TAlignHor);
    procedure FaixaAlignVertical(Faixa: Variant; Alinhamento: TalignVer);
    procedure FaixaBordas(Faixa: Variant; direita, esquerda, superior,inferior: Boolean);
    procedure FaixaValor(Faixa, AValue: Variant);
    procedure FaixaMesclar(Faixa: Variant);
    procedure Cells(Linha,Coluna : Integer; AValue : Variant);
    function cells(Linha,Coluna : Integer): Variant;
    procedure Formula(Linha,Coluna: Integer; Formula : Variant);
    procedure Dataset(MyDataSet: TDataSet; Linha, Coluna: Integer;Titulo: String;OcultarTitulo: Boolean=false );
    procedure IrParaCelula(AValue :Variant);
    procedure imprimir;
    procedure visualizarImpressao;
    procedure VisualizarExcel;
    procedure OcutarExcel;
    function ExcelAtivado: Boolean;
    procedure SalvarComo(NomedoArquivo: Variant);
    procedure Salvar;
    procedure Abrir(NomedoArquivo: Variant);
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
  published
    property nome : Variant read GetNome write SetNome;
    property SheetName : Variant read GetSheetName write SetSheetName;
    property Celula : Variant read FCelula write SetCelula;
end;


procedure register;


implementation


procedure register;
begin
  {$i excelsheets.lrs}    //LResources
  RegisterComponents('Humberto', [TSheets]);
end;


procedure TStylo.SetbackGround(AValue: TColor);
begin
{$IFDEF MSWINDOWS}
if FbackGround=AValue then Exit;
FbackGround:=AValue;
{$ENDIF}
end;

procedure TStylo.SetAlignHor(AValue: TAlignHor);
begin
{$IFDEF MSWINDOWS}
if FAlignHor=AValue then Exit;
FAlignHor:=AValue;
{$ENDIF}
end;

procedure TStylo.SetAlignVer(AValue: TalignVer);
begin
{$IFDEF MSWINDOWS}
if FAlignVer=AValue then Exit;
FAlignVer:=AValue;
{$ENDIF}
end;

procedure TStylo.SetAlturaLinha(AValue: Integer);
begin
{$IFDEF MSWINDOWS}
  if FAlturaLinha=AValue then Exit;
  FAlturaLinha:=AValue;
{$ENDIF}
end;

procedure TStylo.SetColor(AValue: TColor);
begin
{$IFDEF MSWINDOWS}
if FColor=AValue then Exit;
FColor:=AValue;
{$ENDIF}
end;

procedure TStylo.SetFonte(AValue: String);
begin
{$IFDEF MSWINDOWS}
if FFonte=AValue then Exit;
FFonte:=AValue;
{$ENDIF}
end;

procedure TStylo.SetNegrito(AValue: Boolean);
begin
{$IFDEF MSWINDOWS}
if FNegrito=AValue then Exit;
FNegrito:=AValue;
{$ENDIF}
end;

procedure TStylo.SetSize(AValue: Byte);
begin
{$IFDEF MSWINDOWS}
if FSize=AValue then Exit;
FSize:=AValue;
{$ENDIF}
end;

{ TSheets }

function TSheets.GetCelulaValor: Variant;
begin
{$IFDEF MSWINDOWS}
 result := Sheet.Range[FCELULA].value;
{$ENDIF}
end;

function TSheets.GetColunaWidth: Variant;
begin
{$IFDEF MSWINDOWS}
result := Sheet.range[FCelula].ColumnWidth;
{$ENDIF}
end;

function TSheets.GetLinhaHeight: Variant;
begin
{$IFDEF MSWINDOWS}
result := Sheet.range[FCelula].RowHeight;
{$ENDIF}
end;

function TSheets.GetNome: Variant;
begin
{$IFDEF MSWINDOWS}
Result := FNome;
{$ENDIF}
end;

function TSheets.GetSheetName: Variant;
begin
{$IFDEF MSWINDOWS}
result := FSheetName;
{$ENDIF}
end;

procedure TSheets.CelulaAlignHorizontal(AValue: TAlignHor);
begin
{$IFDEF MSWINDOWS}
Case AValue of
   AHCenter  : Sheet.Range[FCelula].HorizontalAlignment := 3;
   AHRight   : Sheet.Range[FCelula].HorizontalAlignment := 4;
   AHLeft    : Sheet.Range[FCelula].HorizontalAlignment := 2;
 end;
{$ENDIF}
end;

procedure TSheets.CelulaAlignVertical(AValue: TalignVer);
begin
{$IFDEF MSWINDOWS}
Case AValue of
   AVTop    : Sheet.Range[FCelula].VerticalAlignment := 1;
   AVCenter : Sheet.Range[FCelula].VerticalAlignment := 2;
   AVBottom : Sheet.Range[FCelula].VerticalAlignment := 3;
end;
{$ENDIF}
end;

procedure TSheets.CelulaBackGround(AValue: TColor);
begin
{$IFDEF MSWINDOWS}
Sheet.range[FCelula].Interior.Color := AValue;
{$ENDIF}
end;

procedure TSheets.CelulaColor(AValue: TColor);
begin
{$IFDEF MSWINDOWS}
Sheet.Range[FCelula].Font.Color := AValue;
{$ENDIF}
end;

procedure TSheets.CelulaFonte(AValue: Variant);
begin
{$IFDEF MSWINDOWS}
Sheet.Range[FCelula].Font.Name := AValue;
{$ENDIF}
end;

procedure TSheets.CelulaMascara(AValue: Variant);
begin
{$IFDEF MSWINDOWS}
Sheet.range[FCELULA].NumberFormatLocal := AValue;
{$ENDIF}
end;

procedure TSheets.CelulaNegrito(AValue: Boolean);
begin
{$IFDEF MSWINDOWS}
Sheet.Range[FCelula].font.Bold := AValue;
{$ENDIF}
end;

procedure TSheets.CelulaSize(AValue: Byte);
begin
{$IFDEF MSWINDOWS}
Sheet.range[FCelula].Font.Size := AValue;
{$ENDIF}
end;

procedure TSheets.CelulaValor(AValue: Variant);
begin
{$IFDEF MSWINDOWS}
Sheet.Range[FCELULA] := AVAlue;
{$ENDIF}
end;

procedure TSheets.CelulaReplace(AValueAtual, AValueNovo: Variant);
begin
{$IFDEF MSWINDOWS}
   Sheet.Range['A1','XFD10000'].Replace(AvalueAtual, AValueNovo);
{$ENDIF}
end;

procedure TSheets.CelulaBordas(direita, esquerda, superior, inferior: Boolean);
begin
{$IFDEF MSWINDOWS}
     if Direita then
        Sheet.Range[FCelula].Borders.Item[$0000000A].Weight  := $00000003;//externo
     if Esquerda then
        Sheet.Range[FCelula].Borders.Item[$00000007].Weight  := $00000003;//externo
     if Superior then
        Sheet.Range[FCelula].Borders.Item[$00000008].Weight   := $00000003;//externo
     if inferior then
        Sheet.Range[FCelula].Borders.Item[$00000009].Weight := $00000003;//externo

{$ENDIF}
end;

procedure TSheets.CelulaSomar(Item1, Item2: Variant);
Var
  Soma : Variant;
begin
{$IFDEF MSWINDOWS}
  Soma := '=SUM('+ Item1 +':'+ Item2 +')';
  Sheet.Range[FCelula].Formula := Soma;
{$ENDIF}
end;

procedure TSheets.CelulaFormatoCurrency(Item1, Item2: Variant);
begin
  Sheet.Cells[Item1,Item2].NumberFormat := 'R$ #,##0.00';
  Sheet.Cells[Item1,Item2].Style := 'Comma';
end;

procedure TSheets.CelulaFormatoTexto(Item1, Item2: Variant);
begin
  Sheet.Cells[Item1,Item2].NumberFormat := ' @';
end;


procedure TSheets.ColunaWidth(AValue: Variant);
begin
{$IFDEF MSWINDOWS}
Sheet.range[FCelula].ColumnWidth := AValue;
{$ENDIF}
end;

function TSheets.ColunaUltima: Integer;
begin
{$IFDEF MSWINDOWS}
Result := Sheet.UsedRange.Columns.Count;
{$ENDIF}
end;

procedure TSheets.ColunaAutoAjuste;
begin
{$IFDEF MSWINDOWS}
Sheet.Columns.AutoFit;
{$ENDIF}
end;

procedure TSheets.ColunaInserir(NumCol: Integer);
begin
{$IFDEF MSWINDOWS}
  ObjExcel.ActiveSheet.Columns[NumCol].Insert;
{$ENDIF}
end;

procedure TSheets.ColunaApagar(NumCol : Integer);
begin
{$IFDEF MSWINDOWS}
  ObjExcel.ActiveSheet.Columns[NumCol].Delete;
{$ENDIF}
end;

procedure TSheets.ColunaSomar(Faixa: Variant);
VAR
    AL : CHAR;
    V,V1 : VARIANT;
    S,N : String;
    i : Byte;
    Letras : array of char;
begin
{$IFDEF MSWINDOWS}
  S := UpperCase(Faixa);
  S := Copy(s,pos(':',s) + 1,Length(s));
  for i := 1 to Length(s) do
     if s[i] in ['A'..'Z'] then
        begin
            SetLength(letras,i );
            Letras[ i -1 ] := s[i] ;
        end
      else
         n := n + s[i];   //pega os numeros
      s := '';
       for i := 0 to High(Letras) do
          s := s + letras[i];
    v := '=SUM(' + FAIXA + ')';
    n := Inttostr(StrToInt(n) + 1);
    v1 := s + n;
  Sheet.range(v1).Formula := V;
{$ENDIF}
end;


procedure TSheets.LinhaSomar(Faixa: Variant);
VAR
    AL : CHAR;
    V,V1 : VARIANT;
    S,N : String;
    i : Byte;
    Letras : array of char;
begin
{$IFDEF MSWINDOWS}
  S := UpperCase(Faixa);
  S := Copy(s,pos(':',s) + 1,Length(s));
  for i := 1 to Length(s) do
     if s[i] in ['A'..'Z'] then
        begin
            SetLength(letras,i );
            Letras[ i -1 ] := s[i] ;
        end
      else
         n := n + s[i];   //pega os numeros
      Case High(Letras) of
         0 :
          if Letras[High(Letras)] = 'Z' then
              begin
                  SetLength(letras,2);
                  Letras[ low(Letras) ] := 'A';
                  Letras[ high(Letras) ] := 'A';
                  n := '1';
              end
              Else
                 Letras[ high(Letras) ] := succ( Letras[high(Letras)] );
         1 :
         if Letras[High(Letras)] = 'Z' then
             begin
                 Letras[ low(Letras) ] := succ( Letras[low(Letras)] );
                 Letras[ high(Letras) ] := 'A';
                 n := '1';
             end
             Else
                Letras[ high(Letras) ] := succ( Letras[high(Letras)] );
      end;
      s := '';
       for i := 0 to High(Letras) do
          s := s + letras[i];
    v := '=SUM(' + FAIXA + ')';
    v1 := s + n;
  Sheet.range(v1).Formula := V;
{$ENDIF}
end;

procedure TSheets.LinhaHeight(AValue: Variant);
begin
{$IFDEF MSWINDOWS}
Sheet.range[FCelula].RowHeight := AValue;
{$ENDIF}
end;

procedure TSheets.LinhaAutoAjuste;
begin
{$IFDEF MSWINDOWS}
Sheet.rows.AutoFit;
{$ENDIF}
end;

procedure TSheets.LinhaApagar(NumLin: Integer);
begin
{$IFDEF MSWINDOWS}
  ObjExcel.ActiveSheet.rows[NumLin].Delete;
{$ENDIF}
end;

function TSheets.LinhaUltima: Integer;
begin
{$IFDEF MSWINDOWS}
  Result := Sheet.UsedRange.Rows.Count;
{$ENDIF}
end;

procedure TSheets.LinhaInserir(NumLin: Integer);
begin
{$IFDEF MSWINDOWS}
  ObjExcel.ActiveSheet.rows[NumLin].Insert;
{$ENDIF}
end;

procedure TSheets.SetCelula(AValue: Variant);
begin
{$IFDEF MSWINDOWS}
if FCelula=AValue then Exit;
FCelula:=AValue;
{$ENDIF}
end;


procedure TSheets.SetNome(AValue: Variant);
begin
{$IFDEF MSWINDOWS}
 if FNome = AValue then  Exit;
 FNome := AValue;
 objExcel.Workbooks[1].WorkSheets[1].Name := AValue;
 objExcel.Caption := AValue;
{$ENDIF}
end;

procedure TSheets.SetSheetName(AValue: Variant);
begin
{$IFDEF MSWINDOWS}
 if FSheetName = AValue then  Exit;
 FSheetName := AValue;
 objExcel.Workbooks[1].WorkSheets[1].Name := AValue;
{$ENDIF}
end;

function TSheets.Alfabetico(qtd : integer) : String;
const
  alpha = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
var
  i,item1,item2,item3 : integer;
begin
  item1 := 0;
  item2 := 0;
  item3 := 0;
  i := 1;
  while  i <= qtd do
          begin
               item1 += 1;
               if item1 = 27 then
                  begin
                    item1 := 1; //ate aqui ok
                    item2 += 1;
                    if item2 = 27 then
                       begin
                         item2 := 1;
                         item3 += 1;
                         if item3 = 27 then
                            item3 := 1;
                       end;
                  end;
               inc(i);
          end;

  if ord(Alpha[item1]) <> 26 then
  result := Alpha[item1] ;
  if ord(Alpha[item2]) <> 26 then
  result := Alpha[item2] + Alpha[item1] ;
  if ord(Alpha[item3]) <> 26 then
  result := Alpha[item3] + Alpha[item2] + Alpha[item1] ;

end;

procedure TSheets.FaixaFonte(Faixa: Variant; FontName: Variant);
begin
{$IFDEF MSWINDOWS}
Sheet.Range[Faixa].font.Name := FontName;
{$ENDIF}
end;

procedure TSheets.FaixaSize(Faixa: Variant; FontSize: byte);
begin
{$IFDEF MSWINDOWS}
Sheet.Range[Faixa].font.size := FontSize;
{$ENDIF}
end;

procedure TSheets.FaixaNegrito(Faixa: Variant; Negrito: Boolean);
begin
{$IFDEF MSWINDOWS}
 Sheet.Range[Faixa].font.Bold := Negrito;
{$ENDIF}
end;

procedure TSheets.FaixaColor(Faixa: Variant; Cor: TColor);
begin
{$IFDEF MSWINDOWS}
 Sheet.Range[Faixa].font.Color := Cor;
{$ENDIF}
end;

procedure TSheets.FaixaBackGround(Faixa: Variant; Cor: TColor);
begin
{$IFDEF MSWINDOWS}
Sheet.Range[Faixa].interior.Color := Cor;
{$ENDIF}
end;

procedure TSheets.FaixaMascara(Faixa: Variant; Mascara: Variant);
begin
{$IFDEF MSWINDOWS}
 Sheet.Range[Faixa].NumberFormat := Mascara;
{$ENDIF}
end;

procedure TSheets.FaixaColunaAutoAjuste(Faixa: Variant);
begin
{$IFDEF MSWINDOWS}
 Sheet.Range[Faixa].columns.AutoFit;
{$ENDIF}
end;

procedure TSheets.FaixaLinhaAutoAjuste(Faixa: Variant);
begin
{$IFDEF MSWINDOWS}
 Sheet.Range[Faixa].rows.AutoFit;
{$ENDIF}
end;

procedure TSheets.FaixaLinhaHeight(Faixa: Variant; height: Variant);
begin
{$IFDEF MSWINDOWS}
 Sheet.Range[Faixa].RowHeight := Height
{$ENDIF}
end;

procedure TSheets.FaixaColunaWidth(Faixa: Variant; Width: Variant);
begin
{$IFDEF MSWINDOWS}
  Sheet.range[Faixa].ColumnWidth := Width;
{$ENDIF}
end;

procedure TSheets.FaixaAlignHorizontal(Faixa: Variant; Alinhamento: TAlignHor);
begin
{$IFDEF MSWINDOWS}
Case Alinhamento of
 AHCenter  : Sheet.Range[Faixa].HorizontalAlignment := 3;
 AHRight   : Sheet.Range[Faixa].HorizontalAlignment := 4;
 AHLeft    : Sheet.Range[Faixa].HorizontalAlignment := 2;
end;
{$ENDIF}
end;

procedure TSheets.FaixaAlignVertical(Faixa: Variant; Alinhamento: TalignVer);
begin
{$IFDEF MSWINDOWS}
Case Alinhamento of
 AVTop    : Sheet.Range[Faixa].VerticalAlignment := 1;
 AVCenter : Sheet.Range[Faixa].VerticalAlignment := 2;
 AVBottom : Sheet.Range[Faixa].VerticalAlignment := 3;
end;
{$ENDIF}
end;

procedure TSheets.FaixaBordas(Faixa: Variant; direita, esquerda, superior,
  inferior: Boolean);
begin
{$IFDEF MSWINDOWS}
     if Direita then
        Sheet.Range[Faixa].Borders.Item[$0000000A].Weight  := $00000003;//externo
     if Esquerda then
        Sheet.Range[Faixa].Borders.Item[$00000007].Weight  := $00000003;//externo
     if Superior then
        Sheet.Range[Faixa].Borders.Item[$00000008].Weight   := $00000003;//externo
     if inferior then
        Sheet.Range[Faixa].Borders.Item[$00000009].Weight := $00000003;//externo
{$ENDIF}
end;

procedure TSheets.FaixaValor(Faixa, AValue: Variant);
begin
{$IFDEF MSWINDOWS}
Sheet.Range[Faixa] := AVAlue;
{$ENDIF}
end;

procedure TSheets.FaixaMesclar(Faixa: Variant);
begin
{$IFDEF MSWINDOWS}
  Sheet.Range[Faixa].MERGE;
{$ENDIF}
end;


procedure TSheets.Cells(Linha, Coluna: Integer; AValue: Variant);
begin
{$IFDEF MSWINDOWS}
Sheet.Cells[Linha,Coluna] := AValue;
{$ENDIF}
end;

function TSheets.cells(Linha, Coluna: Integer): Variant;
begin
{$IFDEF MSWINDOWS}
result := ObjExcel.ActiveSheet.Cells[Linha,Coluna].Value;
{$ENDIF}
end;


procedure TSheets.Formula(Linha, Coluna: Integer; Formula: Variant);
begin
{$IFDEF MSWINDOWS}
Sheet.Cells[Linha,Coluna].NumberFormat := '#.##0,00';
Sheet.Cells[Linha,Coluna].formula := Formula;
{$ENDIF}
end;

procedure TSheets.Dataset(MyDataSet: TDataSet; Linha, Coluna: Integer;
  Titulo: String; OcultarTitulo: Boolean);
var
  Lin,Col : Integer;
  Value : Variant;
  alpha1,alpha2 : variant;
begin
{$IFDEF MSWINDOWS}
if OcultarTitulo = False then
    begin
          for Col := 0 to MyDataSet.FieldCount -1 do     //colunas (titulo)
               begin
                    Value := MyDataSet.Fields.Fields[col].DisplayLabel;
                    Sheet.Cells[Linha,Coluna + Col] := Value;
                    Sheet.Cells[Linha,Coluna + Col].interior.color := clyellow;
                    Sheet.Cells[Linha,Coluna + Col].Font.color := clred;
                    Sheet.Cells[Linha,Coluna + Col].Font.size := 13;
                    Sheet.Cells[Linha,Coluna + Col].HorizontalAlignment := 3;
                    Sheet.Cells[Linha,Coluna + Col].RowHeight := 25;
               end;
          alpha1 := Alfabetico(coluna) + linha.tostring;
          alpha2 := Alfabetico(coluna + col) + linha.ToString;
          IrParaCelula(alpha1);
          LinhaInserir(1);
          FaixaMesclar( alpha1+':'+alpha2);
          FaixaValor(Alpha1+':'+alpha2,Titulo);
          FaixaBackGround(Alpha1+':'+alpha2,clblack);
          FaixaColor(Alpha1+':'+alpha2,clRed);
          FaixaSize(Alpha1+':'+alpha2,14);
          LinhaHeight(40);
          CelulaAlignHorizontal(AHCenter);
          CelulaAlignVertical(AVCenter);
   end;
Linha += 2;
for Lin := 0 to MyDataSet.RecordCount -1 do
   begin
       for Col := 0 to MyDataSet.FieldCount -1 do
            begin

                if MyDataSet.Fields.Fields[col].DataType in [ftDateTime, ftTimeStamp] then
                begin
                     Sheet.Cells[Linha + Lin,Coluna + Col].NumberFormat := 'dd/mm/aaaa hh:mmConfuseds';
                     Value := MyDataSet.Fields.Fields[Col].AsString;
                     Sheet.Cells[Linha + Lin,Coluna + Col]:= Value;
                end
                else
                  if MyDataSet.Fields.Fields[col].DataType in [ftBCD, ftFMTBcd, ftFloat, ftCurrency] then
                  begin
                    Sheet.Cells[Linha + Lin,Coluna + Col].NumberFormat := 'R$ #,##0.00';
                    Sheet.Cells[Linha + Lin,Coluna + Col].Style := 'Comma';
                    Value := MyDataSet.Fields.Fields[Col].AsFloat;
                    Sheet.Cells[Linha + Lin,Coluna + Col]:= Value;
                  end
                  else begin
                    Sheet.Cells[Linha + Lin,Coluna + Col].NumberFormat := ' @';
                    Value := MyDataSet.Fields.Fields[Col].AsString;
                    Sheet.Cells[Linha + Lin,Coluna + Col] := Value;
                  end;

            end;
       MyDataSet.Next;
   end;
 Sheet.Cells[linha + lin,coluna].rows.AutoFit;
 Sheet.rows.AutoFit;
 Sheet.Columns.AutoFit;
{$ENDIF}
end;


procedure TSheets.IrParaCelula(AValue: Variant);
begin
{$IFDEF MSWINDOWS}
  FCelula := AValue;
{$ENDIF}
end;

procedure TSheets.imprimir;
begin
{$IFDEF MSWINDOWS}
Sheet.PrintOut;
{$ENDIF}
end;

procedure TSheets.visualizarImpressao;
begin
{$IFDEF MSWINDOWS}
Sheet.PrintPreview;
{$ENDIF}
end;

procedure TSheets.VisualizarExcel;
begin
{$IFDEF MSWINDOWS}
  objExcel.Visible := True;
  FVisualizaExcel := True;
{$ENDIF}
end;

procedure TSheets.OcutarExcel;
begin
{$IFDEF MSWINDOWS}
   objExcel.Visible := False;
   FVisualizaExcel := False;
{$ENDIF}
end;

function TSheets.ExcelAtivado: Boolean;
begin
{$IFDEF MSWINDOWS}
  Result := FVisualizaExcel;
{$ENDIF}
end;

procedure TSheets.Abrir(NomedoArquivo: Variant);
begin
{$IFDEF MSWINDOWS}
ObjExcel.Quit;
Sheet := Unassigned;
ObjExcel.Workbooks.Open(NomedoArquivo);
Sheet := ObjExcel.ActiveSheet;
if ExcelAtivado then
   VisualizarExcel;
{$ENDIF}
end;


procedure TSheets.SalvarComo(NomedoArquivo: Variant);
begin
{$IFDEF MSWINDOWS}
Sheet.SaveAs(NomedoArquivo);
{$ENDIF}
end;

procedure TSheets.Salvar;
begin
{$IFDEF MSWINDOWS}
ObjExcel.save;
{$ENDIF}
end;

constructor TSheets.Create(AOwner: TComponent);
begin
 Inherited create(AOwner);
{$IFDEF MSWINDOWS}
  Try
     objExcel := CreateOleObject('Excel.Application');
     objExcel.DisplayAlerts := False;
  Except
      raise Exception.Create('Excel está instalado?');
      Exit;
  end;  
  objExcel.Workbooks.Add;
  objExcel.Workbooks[1].Sheets.Add;
  objExcel.Workbooks[1].WorkSheets[1].Name := 'documento';
  Sheet := objExcel.Workbooks[1].WorkSheets['documento'];
  FCelula := 'A1';
{$ENDIF}
end;

destructor TSheets.Destroy;
begin
{$IFDEF MSWINDOWS}
ObjExcel.DisplayAlerts := False;
ObjExcel.Quit;
Sheet:= Unassigned;
objExcel := Unassigned;
inherited Destroy;
{$ENDIF}
end;


end.

