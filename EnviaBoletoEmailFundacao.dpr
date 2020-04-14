program EnviaBoletoEmailFundacao;

{$APPTYPE CONSOLE}

uses
  SysUtils, Classes, IniFiles, Variants, ComObj, Controls, StdCtrls, Windows, CobreBemX_TLB in 'CobreBemX_TLB.pas';
  
var
   mesRef: String;
   anoRef: String;

function FormataData(data: String): TDateTime;
begin
     if AnsiUpperCase(Copy(data, 4, 3)) = 'JAN'
     then
         mesRef := '01'
     else if AnsiUpperCase(Copy(data, 4, 3)) = 'FEB'
     then
         mesRef := '02'
     else if AnsiUpperCase(Copy(data, 4, 3)) = 'MAR'
     then
         mesRef := '03'
     else if AnsiUpperCase(Copy(data, 4, 3)) = 'APR'
     then
         mesRef := '04'
     else if AnsiUpperCase(Copy(data, 4, 3)) = 'MAY'
     then
         mesRef := '05'
     else if AnsiUpperCase(Copy(data, 4, 3)) = 'JUN'
     then
         mesRef := '06'
     else if AnsiUpperCase(Copy(data, 4, 3)) = 'JUL'
     then
         mesRef := '07'
     else if AnsiUpperCase(Copy(data, 4, 3)) = 'AUG'
     then
         mesRef := '08'
     else if AnsiUpperCase(Copy(data, 4, 3)) = 'SEP'
     then
         mesRef := '09'
     else if AnsiUpperCase(Copy(data, 4, 3)) = 'OCT'
     then
         mesRef := '10'
     else if AnsiUpperCase(Copy(data, 4, 3)) = 'NOV'
     then
         mesRef := '11'
     else if AnsiUpperCase(Copy(data, 4, 3)) = 'DEC'
     then
         mesRef := '12';
     anoRef := '20' + Copy(data, 8, 2);
     Result := StrToDate(Copy(data, 1, 2) + '/' + mesRef + '/' + Copy(data, 8, 2));
end;

procedure EnviaBoletosEmail;
var
   arqBoletos: TStringList;
   linhaBoleto: TStringList;
   sr: TSearchRec;
   localizacao: String;
   arquivoIni: TIniFile;
   caminhoArqs: String;
   extensaoArqs: String;
   URLImagens: String;
   caminhoLicenca: String;
   caminhoHTMLReciboPersonalizado: String;
   HTMLReciboPersonalizado: TStringList;
   agencia: String;
   contaCorrente: String;
   codigoCedente: String;
   varCarteira: String;
   instCaixa: String;
   localPagamento: String;
   servidor: String;
   porta: Integer;
   usuario: String;
   senha: String;
   assunto: String;
   enderecoEmail: String;
   nome: String;
   CobreBemX: IContaCorrente;
   Boleto: Variant;
   EmailSacado: Variant;
   i: Integer;
   qtdBoletos: Integer;
begin
     localizacao := IncludeTrailingPathDelimiter(ExtractFilePath(ParamStr(0)));
     arquivoIni := TIniFile.Create(ChangeFileExt(ParamStr(0), '.ini'));
//CONFIGURACOES DA APLICACAO
     caminhoArqs := arquivoIni.ReadString('Parametros', 'caminhoArquivos', '');
     extensaoArqs := arquivoIni.ReadString('Parametros', 'extensaoArquivos', '');
//CONFIGURACOES DO BOLETO
     URLImagens := arquivoIni.ReadString('ConfigsCBX', 'URLImagens', '');
     caminhoLicenca := arquivoIni.ReadString('ConfigsCBX', 'caminhoLicenca', '');
     caminhoHTMLReciboPersonalizado := arquivoIni.ReadString('ConfigsCBX', 'caminhoHTMLReciboPersonalizado', '');
     HTMLReciboPersonalizado := TStringList.Create;
     HTMLReciboPersonalizado.LoadFromFile(caminhoHTMLReciboPersonalizado);
     agencia := arquivoIni.ReadString('ConfigsCBX', 'agencia', '');
     contaCorrente := arquivoIni.ReadString('ConfigsCBX', 'contaCorrente', '');
     codigoCedente := arquivoIni.ReadString('ConfigsCBX', 'codigoCedente', '');
     varCarteira := arquivoIni.ReadString('ConfigsCBX', 'varCarteira', '');
     instCaixa := arquivoIni.ReadString('ConfigsCBX', 'instCaixa', '');
     localPagamento := arquivoIni.ReadString('ConfigsCBX', 'localPagamento', '');
//CONFIGURACOES DO SERVIDOR DE EMAIL
     servidor := arquivoIni.ReadString('ServidorEmail', 'servidor', '');
     porta := StrToInt(arquivoIni.ReadString('ServidorEmail', 'porta', ''));
     usuario := arquivoIni.ReadString('ServidorEmail', 'usuario', '');
     senha := arquivoIni.ReadString('ServidorEmail', 'senha', '');
//CONFIGURACOES DO ENVIO DE EMAIL
     assunto := arquivoIni.ReadString('DadosEmail', 'assunto', '');
     enderecoEmail := arquivoIni.ReadString('DadosEmail', 'enderecoEmail', '');
     nome := arquivoIni.ReadString('DadosEmail', 'nome', '');

     arquivoINI.Free;

     WriteLn('INICIO - INSTANCIANDO O CBX.');
     CoInitializeEx(nil, 0);
     CobreBemX := CoContaCorrente.Create;     CobreBemX.ArquivoLicenca := caminhoLicenca;

     if CobreBemX.UltimaMensagemErro <> ''
     then
         begin
              WriteLn('ERRO NO CBX COM A MENSAGEM: ' + CobreBemX.UltimaMensagemErro);
              WriteLn('APERTE ALGUMA TECLA PARA FINALIZAR A APLICACAO.');
              ReadLn;
         end;
     WriteLn('FIM - INSTANCIANDO CBX.');

     CobreBemX.CodigoAgencia := agencia;
     CobreBemX.NumeroContaCorrente := contaCorrente;
     CobreBemX.CodigoCedente := codigoCedente;
     CobreBemX.OutroDadoConfiguracao1 := varCarteira;
     CobreBemX.InicioNossoNumero := '00000000001';
     CobreBemX.FimNossoNumero := '99999999999';
     CobreBemX.ProximoNossoNumero := '1';
     CobreBemX.LocalPagamento := localPagamento;

     CobreBemX.PadroesBoleto.PadroesBoletoEmail.SMTP.Servidor := servidor;
     CobreBemX.PadroesBoleto.PadroesBoletoEmail.SMTP.Porta := porta;
     CobreBemX.PadroesBoleto.PadroesBoletoEmail.SMTP.Usuario := usuario;
     CobreBemX.PadroesBoleto.PadroesBoletoEmail.SMTP.Senha := senha;

     CobreBemX.PadroesBoleto.PadroesBoletoEmail.URLImagensCodigoBarras := URLImagens;
     CobreBemX.PadroesBoleto.PadroesBoletoEmail.PadroesEmail.Assunto := assunto;
     CobreBemX.PadroesBoleto.PadroesBoletoEmail.PadroesEmail.EmailFrom.Endereco := enderecoEmail;
     CobreBemX.PadroesBoleto.PadroesBoletoEmail.PadroesEmail.EmailFrom.Nome := nome;
     CobreBemX.PadroesBoleto.PadroesBoletoEmail.LayoutBoletoEmail := 'PadraoReciboPersonalizadoImpressao';
     CobreBemX.PadroesBoleto.PadroesBoletoImpresso.HTMLReciboPersonalizado := HTMLReciboPersonalizado.Text;

     try
        if FindFirst(caminhoArqs + '*.*', faArchive, sr) = 0
        then
            begin
                 repeat
                       if (AnsiUpperCase(ExtractFileExt(sr.Name)) = AnsiUpperCase(extensaoArqs))
                       then
                           begin
                                arqBoletos := TStringList.Create;
                                arqBoletos.LoadFromFile(caminhoArqs + sr.Name);

                                linhaBoleto := TStringList.Create;
                                linhaBoleto.Delimiter := ';';
                                linhaBoleto.StrictDelimiter := True;

                                CobreBemX.DocumentosCobranca.Clear;

                                WriteLn('INICIO - PROCESSANDO ARQUIVO: ' + sr.Name);
                                qtdBoletos := 0;
                                for i := 1 to arqBoletos.Count - 1 do
                                begin
                                     qtdBoletos := qtdBoletos + 1;
                                     linhaBoleto.Clear;
                                     linhaBoleto.DelimitedText := arqBoletos.Strings[i];

                                     Boleto := CobreBemX.DocumentosCobranca.Add;
                                     Boleto.NumeroDocumento := linhaBoleto[15] + '/' + linhaBoleto[4];
                                     Boleto.Aceite := 'N';
                                     Boleto.NossoNumero := codigoCedente + Format('%.11d', [StrToInt(linhaBoleto[16])]);
                                     Boleto.DataDocumento := FormatDateTime('dd/mm/yyyy', StrToDate(linhaBoleto[31]));
                                     Boleto.DataVencimento  := FormatDateTime('dd/mm/yyyy', FormataData(linhaBoleto[17]));
                                     Boleto.ValorDocumento := 0.00;
                                     instCaixa := StringReplace(instCaixa, '<#MM>', mesRef, [rfReplaceAll]);
                                     instCaixa := StringReplace(instCaixa, '<#AAAA>', anoRef, [rfReplaceAll]);
                                     Boleto.PadroesBoleto.InstrucoesCaixa := instCaixa;
                                     Boleto.NomeSacado := linhaBoleto[10];

                                     EmailSacado := Boleto.EnderecosEmailSacado.Add;
                                     EmailSacado.Nome := Boleto.NomeSacado;
                                     EmailSacado.Endereco := linhaBoleto[30];
                                end;
                                WriteLn('FIM - PROCESSANDO ARQUIVO: ' + sr.Name);
                                WriteLn('INICIO - ENVIANDO ' + IntToStr(qtdBoletos) + ' BOLETOS POR EMAIL.');
                                CobreBemX.EnviaBoletosPorEmail;
                                WriteLn('FIM - ENVIANDO ' + IntToStr(qtdBoletos) + ' BOLETOS POR EMAIL.');
                           end;
                 until FindNext(sr) <> 0;
                       SysUtils.FindClose(sr);
            end;
     finally
            arqBoletos.Free;
            linhaBoleto.Free;
     end;
end;

begin
     WriteLn('INICIO - PROCESSA E ENVIA BOLETOS EMAIL.');
     EnviaBoletosEmail;
     WriteLn('FIM - PROCESSA E ENVIA BOLETOS EMAIL.');
     ReadLn;
end.
