/**
 * Projeto: Gerador Automático de Recibos
 * Tecnologias: Google Apps Script + Google Sheets + Google Docs
 * Descrição: Automatiza a geração de recibos a partir de planilhas,
 * organizando arquivos por período, entidade e quinzena.
 */

// -----------------------------------------------------
// CONFIGURAÇÕES (SEM DADOS REAIS)
// -----------------------------------------------------

const CONFIG = {
  ROOT_FOLDER: 'Billing',
  RECEIPTS_FOLDER: 'Receipts',
  ENTITIES: {
    ENTITY_A: {
      name: 'HEALTHCARE ENTITY A',
      logo: 'ENTITY_A_LOGO.png',
      bankInfo: [
        'BANK: XXXXX',
        'AGENCY: XXXX-X',
        'ACCOUNT: XXXXX-X'
      ]
    },
    ENTITY_B: {
      name: 'HEALTHCARE ENTITY B',
      logo: 'ENTITY_B_LOGO.png',
      bankInfo: [
        'BANK: XXXXX',
        'AGENCY: XXXX-X',
        'ACCOUNT: XXXXX-X'
      ]
    }
  }
};

// -----------------------------------------------------
// FUNÇÃO PRINCIPAL
// -----------------------------------------------------

function generateReceipts() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const periodName = sheet.getName(); // Ex: january-2025

  const baseFolder = getOrCreateFolder(null, CONFIG.ROOT_FOLDER);
  const receiptsFolder = getOrCreateFolder(baseFolder, CONFIG.RECEIPTS_FOLDER);
  const periodFolder = getOrCreateFolder(receiptsFolder, periodName);

  data.forEach((row, index) => {
    if (index === 0) return; // header

    const [description, summary, amount, invoice] = row;
    if (!invoice || !amount) return;

    const entityKey = description.includes('A') ? 'ENTITY_A' : 'ENTITY_B';
    const entity = CONFIG.ENTITIES[entityKey];

    const half = index < 20 ? 'First_Half' : 'Second_Half';
    const entityFolder = getOrCreateFolder(periodFolder, entity.name);
    const halfFolder = getOrCreateFolder(entityFolder, half);

    if (fileExists(halfFolder, `Receipt - ${invoice}`)) return;

    createReceiptDocument(entity, summary, amount, invoice, halfFolder);
  });
}

// -----------------------------------------------------
// DOCUMENTO
// -----------------------------------------------------

function createReceiptDocument(entity, summary, amount, invoice, destinationFolder) {
  const doc = DocumentApp.create(`Receipt - ${invoice}`);
  const body = doc.getBody();

  body.appendParagraph(entity.name).setBold(true).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph('RECEIPT').setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body.appendParagraph(
    `Received the amount of $ ${formatCurrency(amount)} referring to ${summary}.
    Invoice: ${invoice}`
  ).setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);

  entity.bankInfo.forEach(line => body.appendParagraph(line));

  body.appendParagraph('_____________________________');
  body.appendParagraph('Authorized Signature');

  doc.saveAndClose();

  const file = DriveApp.getFileById(doc.getId());
  destinationFolder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);
}

// -----------------------------------------------------
// UTILITÁRIOS
// -----------------------------------------------------

function getOrCreateFolder(parent, name) {
  const folders = parent ? parent.getFoldersByName(name) : DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : (parent || DriveApp.getRootFolder()).createFolder(name);
}

function fileExists(folder, name) {
  return folder.getFilesByName(name).hasNext();
}

function formatCurrency(value) {
  return Number(value).toLocaleString('en-US', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}
