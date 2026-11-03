/**
 * Organizador de Contas Atrasadas — trava as fórmulas no Google Sheets.
 *
 * Como usar (uma vez só, na SUA planilha original, antes de gerar o link /copy):
 *   1. Abra a planilha no Google Sheets.
 *   2. Menu Extensões › Apps Script.
 *   3. Apague o que estiver lá, cole este arquivo inteiro e clique em Salvar (💾).
 *   4. No topo, escolha a função "protegerPlanilha" e clique em Executar.
 *   5. Autorize o acesso (Avançado › Acessar projeto) — o script mexe só nesta planilha.
 *
 * Resultado: em todas as abas só as células de preenchimento ficam livres.
 * Quem tentar editar uma fórmula vê um aviso. As proteções vão junto quando
 * o comprador faz a cópia pelo link /copy.
 * Para tirar tudo de novo, rode a função "removerProtecoes".
 */

var ENTRADAS = {
  'Painel': ['B6:C6', 'E6:F6', 'H6:I6'],
  'Contas': ['C6:E55', 'G6:H55', 'L6:L55', 'N6:O55'],
  'Pagamentos': ['B6:F505'],
  'Plano de Quitação': ['D5:F5', 'J4:K5'],
  'Como usar': []
};

function protegerPlanilha() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  removerProtecoes();

  Object.keys(ENTRADAS).forEach(function (nome) {
    var aba = ss.getSheetByName(nome);
    if (!aba) return;
    var protecao = aba.protect().setDescription('Fórmulas do Organizador — ' + nome);
    var livres = ENTRADAS[nome].map(function (a1) { return aba.getRange(a1); });
    if (livres.length) protecao.setUnprotectedRanges(livres);
    // Aviso em vez de bloqueio total: funciona para quem fizer a cópia
    // e evita que a pessoa fique "presa" sem conseguir usar.
    protecao.setWarningOnly(true);
  });

  var listas = ss.getSheetByName('Listas');
  if (listas) {
    listas.hideSheet();
    listas.protect().setDescription('Listas dos menus').setWarningOnly(true);
  }

  ss.setActiveSheet(ss.getSheetByName('Painel'));
  SpreadsheetApp.getUi().alert('Pronto! As fórmulas estão protegidas. Só as células de preenchimento ficam livres.');
}

function removerProtecoes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheets().forEach(function (aba) {
    aba.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function (p) { p.remove(); });
    aba.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function (p) { p.remove(); });
  });
}
