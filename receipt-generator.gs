function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📄 Gerador de Recibos')
    .addItem('📅 Gerar Recibos', 'mostrarDialogo')
    .addItem('📂 Abrir Pasta', 'abrirPasta')
    .addToUi();
}

function mostrarDialogo() {
  const html = HtmlService.createHtmlOutput(`
    <html>
      <body style="font-family: sans-serif; padding: 10px;">
        <p>Preencha <b>um</b> dos campos abaixo para gerar os recibos:</p>
        <div style="margin-bottom: 15px;">
          <label>Número da Nota (Coluna D):</label><br>
          <input type="text" id="nota" style="width: 100%; padding: 5px; margin-top: 5px;">
        </div>
        <div style="margin-bottom: 15px;">
          <label>Data do Protocolo (Coluna E):</label><br>
          <input type="text" id="data" placeholder="ex: 01/01/2026" style="width: 100%; padding: 5px; margin-top: 5px;">
        </div>
        <button onclick="enviar()" style="width: 100%; padding: 10px; background-color: #4285f4; color: white; border: none; border-radius: 4px; cursor: pointer;">Gerar Recibos</button>
        
        <script>
          function enviar() {
            const nota = document.getElementById('nota').value;
            const data = document.getElementById('data').value;
            if (!nota && !data) {
              alert('Por favor, preencha ao menos um dos campos.');
              return;
            }
            google.script.run.withSuccessHandler(() => google.script.host.close()).gerarDocs(nota, data);
          }
        </script>
      </body>
    </html>
  `)
  .setWidth(350)
  .setHeight(250);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Critérios de Geração');
}

function abrirPasta() {
  const ss = SpreadsheetApp.getActive();
  const nomePlanilha = ss.getActiveSheet().getName();
  const ui = SpreadsheetApp.getUi();
  
  try {
    const pastaRaiz = DriveApp.getFoldersByName("ARQUIVOS_GERAIS").next();
    const pastaSub = pastaRaiz.getFoldersByName("RECIBOS_GERADOS").next();
    const pastaDestino = pastaSub.getFoldersByName(nomePlanilha).next();
    
    const url = pastaDestino.getUrl();
    const htmlOutput = HtmlService
      .createHtmlOutput(`<p>Clique no botão abaixo para acessar a pasta:</p>
                         <a href="${url}" target="_blank" 
                            style="padding: 10px 20px; background-color: #4285f4; color: white; text-decoration: none; border-radius: 5px; font-family: sans-serif; display: inline-block;">
                            Abrir Pasta no Drive
                         </a>`)
      .setWidth(350)
      .setHeight(150);
    
    ui.showModalDialog(htmlOutput, 'Pasta Localizada');
    
  } catch (e) {
    ui.alert("A pasta ainda não foi localizada. Gere um recibo primeiro.");
  }
}

function gerarDocs(notaDigitada, dataDigitada) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  const nomePlanilha = sheet.getName();
  const ultimaLinha = sheet.getLastRow();
  const dados = sheet.getRange(1, 1, ultimaLinha, 5).getValues();

  notaDigitada = notaDigitada ? notaDigitada.trim() : "";
  dataDigitada = dataDigitada ? dataDigitada.trim() : "";

  let pastaRaiz;
  const buscaRaiz = DriveApp.getFoldersByName("ARQUIVOS_GERAIS");
  if (buscaRaiz.hasNext()) {
    pastaRaiz = buscaRaiz.next();
  } else {
    ui.alert("Pasta principal não encontrada!");
    return;
  }

  const buscaSub = pastaRaiz.getFoldersByName("RECIBOS_GERADOS");
  let pastaSub = buscaSub.hasNext() ? buscaSub.next() : pastaRaiz.createFolder("RECIBOS_GERADOS");

  const buscaPastaPlanilha = pastaSub.getFoldersByName(nomePlanilha);
  let pastaPrincipal = buscaPastaPlanilha.hasNext() ? buscaPastaPlanilha.next() : pastaSub.createFolder(nomePlanilha);

  let pastaA, pastaB;
  const buscaA = pastaPrincipal.getFoldersByName("EMPRESA_A");
  pastaA = buscaA.hasNext() ? buscaA.next() : pastaPrincipal.createFolder("EMPRESA_A");
  const buscaB = pastaPrincipal.getFoldersByName("EMPRESA_B");
  pastaB = buscaB.hasNext() ? buscaB.next() : pastaPrincipal.createFolder("EMPRESA_B");

  const logos = buscarLogos();
  const EMPRESA_1 = 'NOME DA EMPRESA A';
  const EMPRESA_2 = 'NOME DA EMPRESA B';
  
  let empresaAtual = '';
  let contagemA = 0;
  let contagemB = 0;

  dados.forEach((linha, index) => {
    const linhaNumero = index + 1;
    const descricao = linha[0];
    const resumo = linha[1];
    const valor = linha[2];
    const notaFiscal = linha[3] ? linha[3].toString().trim() : "";
    const dataNaPlanilha = linha[4];

    if (linhaNumero === 2 && descricao === EMPRESA_1) { empresaAtual = EMPRESA_1; return; }
    if (linhaNumero === 45 && descricao === EMPRESA_2) { empresaAtual = EMPRESA_2; return; }
    if (!empresaAtual) return;

    let dataFormatada = "";
    if (dataNaPlanilha instanceof Date) {
      dataFormatada = Utilities.formatDate(dataNaPlanilha, Session.getScriptTimeZone(), "dd/MM/yyyy");
    } else {
      dataFormatada = dataNaPlanilha ? dataNaPlanilha.toString().trim() : "";
    }

    let atendeCondicao = false;
    if (notaDigitada !== "" && notaFiscal === notaDigitada) atendeCondicao = true;
    if (dataDigitada !== "" && dataFormatada === dataDigitada) atendeCondicao = true;

    if (!atendeCondicao) return;
    if (!notaFiscal || !resumo || resumo.toString().trim() === '' || !descricao || !valor) return;
    
    // Filtros de intervalo
    if (empresaAtual === EMPRESA_1 && (linhaNumero < 3 || linhaNumero >= 45)) return;
    if (empresaAtual === EMPRESA_2 && linhaNumero < 45) return;

    let pastaDestino = (empresaAtual === EMPRESA_2) ? pastaB : pastaA;
    let sigla = (empresaAtual === EMPRESA_2) ? 'EB' : 'EA';
    let nomeArquivo = `Recibo - ${sigla} - NF ${notaFiscal}`;
    let contador = 1;

    while (pastaDestino.getFilesByName(nomeArquivo).hasNext()) {
      nomeArquivo = `Recibo - ${sigla} - NF ${notaFiscal} (${contador})`;
      contador++;
    }

    const doc = DocumentApp.create(nomeArquivo);
    const corpo = doc.getBody();
    corpo.setMarginTop(30); corpo.setMarginBottom(30); corpo.setMarginLeft(50); corpo.setMarginRight(50);

    if (empresaAtual === EMPRESA_2) {
      criarLayoutB(corpo, descricao, resumo, valor, notaFiscal, logos[empresaAtual]);
      contagemB++;
    } else {
      criarLayoutA(corpo, descricao, resumo, valor, notaFiscal, logos[empresaAtual]);
      contagemA++;
    }

    doc.saveAndClose();
    const arquivo = DriveApp.getFileById(doc.getId());
    pastaDestino.addFile(arquivo);
    DriveApp.getRootFolder().removeFile(arquivo); 
  });

  if ((contagemA + contagemB) > 0) {
    abrirPasta();
  } else {
    ui.alert("Nenhum recibo gerado para os critérios informados.");
  }
}


function criarLayoutA(corpo, descricao, resumo, valor, notaFiscal, logo) {
  let banco, ag, cc;
  const resumoTexto = resumo.toString().toLowerCase();

  if (resumoTexto.includes("tipo_1")) {
    banco = "BANCO EXEMPLO A"; ag = "0000"; cc = "00000-0";
  } else {
    banco = "BANCO EXEMPLO B"; ag = "0000"; cc = "00000-0";
  }

  inserirLogo(corpo, logo);
  corpo.appendParagraph('NOME DA EMPRESA A').setAlignment(DocumentApp.HorizontalAlignment.CENTER).setBold(true);
  corpo.appendParagraph('CNPJ: 00.000.000/0000-00').setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingAfter(20);
  corpo.appendParagraph('RECIBO').setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingAfter(8);
  corpo.appendParagraph(descricao).setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingAfter(15);

  const textoRecibo = `RECEBEMOS DE [NOME DO PAGADOR] A IMPORTÂNCIA DE R$ ${formatarValor(valor)} (${converterValorParaExtenso(valor)}) REFERENTE A ${resumo} - CONFORME NF ${notaFiscal}`;
  corpo.appendParagraph(textoRecibo).setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY).setLineSpacing(1.3).setSpacingAfter(20);

  corpo.appendParagraph('DADOS BANCÁRIOS:').setBold(true).setSpacingAfter(5);
  corpo.appendParagraph(banco).setSpacingAfter(2);
  corpo.appendParagraph(`AGÊNCIA: ${ag}`).setSpacingAfter(2);
  corpo.appendParagraph(`CONTA: ${cc}`).setSpacingAfter(30);

  adicionarAssinatura(corpo);
}

function criarLayoutB(corpo, descricao, resumo, valor, notaFiscal, logo) {
  let banco, ag, cc;
  
  banco = "BANCO EXEMPLO C"; ag = "0000"; cc = "00000-0";

  inserirLogo(corpo, logo);
  corpo.appendParagraph('NOME DA EMPRESA B').setAlignment(DocumentApp.HorizontalAlignment.CENTER).setBold(true);
  corpo.appendParagraph('CNPJ: 00.000.000/0000-00').setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingAfter(20);
  corpo.appendParagraph('RECIBO').setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingAfter(8);
  corpo.appendParagraph(descricao).setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingAfter(15);

  const textoRecibo = `RECEBEMOS DE [NOME DO PAGADOR] A IMPORTÂNCIA DE R$ ${formatarValor(valor)} (${converterValorParaExtenso(valor)}) REFERENTE A ${resumo} - CONFORME NF ${notaFiscal}`;
  corpo.appendParagraph(textoRecibo).setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY).setLineSpacing(1.3).setSpacingAfter(20);

  corpo.appendParagraph('DADOS BANCÁRIOS:').setBold(true).setSpacingAfter(5);
  corpo.appendParagraph(banco).setSpacingAfter(2);
  corpo.appendParagraph(`AGÊNCIA: ${ag}`).setSpacingAfter(2);
  corpo.appendParagraph(`CONTA: ${cc}`).setSpacingAfter(30);

  adicionarAssinatura(corpo);
}

function buscarLogos() {
  const logos = {};
  try {
    const pastas = DriveApp.getFoldersByName("Imagens_Sistema");
    if (pastas.hasNext()) {
      const arquivos = pastas.next().getFiles();
      while (arquivos.hasNext()) {
        const arq = arquivos.next();
        if (arq.getName() === 'LOGO_A.png') logos['NOME DA EMPRESA A'] = arq;
        if (arq.getName() === 'LOGO_B.png') logos['NOME DA EMPRESA B'] = arq;
      }
    }
  } catch (e) { console.log("Erro ao carregar logos: " + e); }
  return logos;
}

function inserirLogo(corpo, logo) {
  if (logo) {
    try {
      const blob = logo.getBlob();
      const img = corpo.appendParagraph('').appendInlineImage(blob);
      const MAX_WIDTH = 200;
      const MAX_HEIGHT = 100;
      let newWidth = img.getWidth();
      let newHeight = img.getHeight();
      const ratio = newWidth / newHeight;

      if (newWidth > MAX_WIDTH) {
        newWidth = MAX_WIDTH;
        newHeight = newWidth / ratio;
      }
      if (newHeight > MAX_HEIGHT) {
        newHeight = MAX_HEIGHT;
        newWidth = newHeight * ratio;
      }
      img.setWidth(newWidth).setHeight(newHeight);
      corpo.getParagraphs()[corpo.getParagraphs().length-1].setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    } catch(e) {}
  }
}

function adicionarAssinatura(corpo) {
  corpo.appendParagraph('').setSpacingBefore(30);
  corpo.appendParagraph('_______________________________________________________').setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  corpo.appendParagraph('Assinatura do Responsável').setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  corpo.appendParagraph('CIDADE - ESTADO').setAlignment(DocumentApp.HorizontalAlignment.CENTER);
}

function formatarValor(valor) {
  let n = typeof valor === 'string' ? parseFloat(valor.replace('R$', '').replace(/\./g, '').replace(',', '.')) : valor;
  return isNaN(n) ? '0,00' : n.toLocaleString('pt-BR', { minimumFractionDigits: 2 });
}

function converterValorParaExtenso(valor) {
  let num = typeof valor === 'string' ? parseFloat(valor.replace('R$', '').replace(/\./g, '').replace(',', '.')) : valor;
  if (isNaN(num)) return "VALOR INVÁLIDO";
  const partes = num.toFixed(2).split('.');
  const inteiro = parseInt(partes[0]);
  const decimal = parseInt(partes[1]);
  let res = (inteiro === 0) ? 'ZERO' : converterCentenas(inteiro);
  res += (inteiro === 1) ? ' REAL' : ' REAIS';
  if (decimal > 0) res += ' E ' + converterCentenas(decimal) + (decimal === 1 ? ' CENTAVO' : ' CENTAVOS');
  return res;
}

function converterCentenas(n) {
  const unidades = ['', 'UM', 'DOIS', 'TRÊS', 'QUATRO', 'CINCO', 'SEIS', 'SETE', 'OITO', 'NOVE'];
  const de10a19 = ['DEZ', 'ONZE', 'DOZE', 'TREZE', 'CATORZE', 'QUINZE', 'DEZESSEIS', 'DEZESSETE', 'DEZOITO', 'DEZENOVE'];
  const dezenas = ['', 'DEZ', 'VINTE', 'TRINTA', 'QUARENTA', 'CINQUENTA', 'SESSENTA', 'SETENTA', 'OITENTA', 'NOVENTA'];
  const centenas = ['', 'CENTO', 'DUZENTOS', 'TREZENTOS', 'QUATROCENTOS', 'QUINHENTOS', 'SEISCENTOS', 'SETECENTOS', 'OITOCENTOS', 'NOVECENTOS'];
  if (n === 0) return '';
  if (n === 100) return 'CEM';
  if (n < 1000) {
    let c = Math.floor(n / 100), d = Math.floor((n % 100) / 10), u = n % 10;
    let r = centenas[c];
    if (d > 0) { r += (r ? ' E ' : '') + (d === 1 ? de10a19[u] : dezenas[d] + (u > 0 ? ' E ' + unidades[u] : '')); }
    else if (u > 0) { r += (r ? ' E ' : '') + unidades[u]; }
    return r;
  }
  let mil = Math.floor(n / 1000), resto = n % 1000;
  let r = (mil === 1 ? 'MIL' : converterCentenas(mil) + ' MIL');
  if (resto > 0) r += (resto < 100 ? ' E ' : ', ') + converterCentenas(resto);
  return r;
}
