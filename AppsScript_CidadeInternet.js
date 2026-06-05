// ════════════════════════════════════════════════════════════════
//  CIDADE DA INTERNET — Apps Script para Google Sheets
//  Cole este código em: Extensões > Apps Script > Novo script
//  Depois publique como Web App (acesso: qualquer pessoa)
// ════════════════════════════════════════════════════════════════

const SHEET_NAME = 'Respostas';

// Colunas do cabeçalho
const HEADERS = [
  'Data/Hora Envio',
  'Nome do Aluno',
  'Turma',
  'Total Concluídas',
  'Finalizado?',
  '📡 Wi-Fi',        'Horário Wi-Fi',
  '🚪 DHCP',         'Horário DHCP',
  '📒 DNS',          'Horário DNS',
  '📦 TCP/IP',       'Horário TCP/IP',
  '🍔 HTTP',         'Horário HTTP',
  '🚚 FTP',          'Horário FTP',
  '📬 SMTP',         'Horário SMTP',
  '🎮 SSH',          'Horário SSH',
];

const PROTO_ORDER = ['wifi','dhcp','dns','tcp','http','ftp','smtp','ssh'];

// ── Recebe POST do jogo HTML ─────────────────────────────────────
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);

    // Cria a aba e o cabeçalho se não existir
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(HEADERS);
      formatHeader(sheet);
    }

    // Garante cabeçalho mesmo se a aba existir vazia
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      formatHeader(sheet);
    }

    // Monta a linha de dados
    const progrMap = {};
    (data.progresso || []).forEach(p => { progrMap[p.protocolo] = p; });

    const row = [
      formatDateBR(data.dataEnvio),    // Data/Hora Envio
      data.nome || '',                  // Nome
      data.turma || '',                 // Turma
      data.total || 0,                  // Total concluídas
      data.finalizado ? 'SIM' : 'NÃO', // Finalizado?
    ];

    PROTO_ORDER.forEach(k => {
      const p = progrMap[k] || {};
      row.push(p.concluido ? '✅' : '❌');
      row.push(p.horario ? formatDateBR(p.horario) : '—');
    });

    // Verifica se já existe linha do mesmo aluno+turma
    const existing = findRow(sheet, data.nome, data.turma);
    if (existing > 0) {
      // Atualiza linha existente
      sheet.getRange(existing, 1, 1, row.length).setValues([row]);
    } else {
      // Insere nova linha
      sheet.appendRow(row);
    }

    // Aplica formatação condicional básica
    applyConditionalFormat(sheet);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', msg: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Também aceita GET para teste rápido ─────────────────────────
function doGet(e) {
  return ContentService
    .createTextOutput('✅ Apps Script da Cidade da Internet está funcionando!')
    .setMimeType(ContentService.MimeType.TEXT);
}

// ── Busca linha existente pelo nome+turma ────────────────────────
function findRow(sheet, nome, turma) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const nomes  = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  const turmas = sheet.getRange(2, 3, lastRow - 1, 1).getValues();
  for (let i = 0; i < nomes.length; i++) {
    if (
      nomes[i][0].toString().trim().toLowerCase() === nome.toString().trim().toLowerCase() &&
      turmas[i][0].toString().trim().toLowerCase() === turma.toString().trim().toLowerCase()
    ) {
      return i + 2; // linha real na planilha (1-indexed, +1 pelo cabeçalho)
    }
  }
  return -1;
}

// ── Formata cabeçalho ────────────────────────────────────────────
function formatHeader(sheet) {
  const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
  headerRange
    .setBackground('#1a1a2e')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 150); // Data
  sheet.setColumnWidth(2, 200); // Nome
  sheet.setColumnWidth(3, 80);  // Turma
  sheet.setColumnWidth(4, 60);  // Total
  sheet.setColumnWidth(5, 80);  // Finalizado
  // Colunas de protocolo e horário
  for (let i = 6; i <= HEADERS.length; i++) {
    sheet.setColumnWidth(i, i % 2 === 0 ? 150 : 50);
  }
}

// ── Formatação condicional (verde = finalizado) ──────────────────
function applyConditionalFormat(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const range = sheet.getRange(2, 5, lastRow - 1, 1); // coluna "Finalizado?"
  const rules = sheet.getConditionalFormatRules();
  // Evita duplicar regras
  const alreadyHasRule = rules.some(r => {
    const ranges = r.getRanges();
    return ranges.some(rng => rng.getA1Notation() === range.getA1Notation());
  });
  if (alreadyHasRule) return;
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('SIM')
    .setBackground('#e6f4ea')
    .setFontColor('#1e8e3e')
    .setRanges([range])
    .build();
  rules.push(rule);
  sheet.setConditionalFormatRules(rules);
}

// ── Converte ISO para data BR legível ────────────────────────────
function formatDateBR(isoStr) {
  if (!isoStr) return '';
  try {
    const d = new Date(isoStr);
    const pad = n => String(n).padStart(2, '0');
    return `${pad(d.getDate())}/${pad(d.getMonth()+1)}/${d.getFullYear()} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
  } catch(e) {
    return isoStr;
  }
}
