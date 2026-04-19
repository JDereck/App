// =====================================================================
// BACKEND — Google Apps Script
// Como usar:
//   1. Abra sua planilha → Extensões → Apps Script
//   2. Apague tudo e cole este código
//   3. Salve (Ctrl+S)
//   4. Implantar → Nova implantação → Aplicativo da Web
//      - Executar como: Eu
//      - Quem tem acesso: Qualquer pessoa
//   5. Copie a URL gerada e cole no index.html
// =====================================================================

var NOME_ABA = 'Tarefas';
var CABECALHOS = ['ID', 'Tarefa', 'Descrição', 'Prioridade', 'Subtarefas', 'Concluído'];

// Busca todas as tarefas e retorna como JSON (método GET)
function doGet(e) {
  var planilha = getAba();
  var dados = planilha.getDataRange().getValues();

  // Pula a linha de cabeçalho (índice 0)
  var linhas = dados.slice(1);

  var tarefas = linhas.map(function(linha) {
    return {
      id:         String(linha[0]),
      tarefa:     linha[1],
      descricao:  linha[2],
      prioridade: linha[3],
      subtarefas: parsearJSON(linha[4]),
      concluido:  linha[5] === true || linha[5] === 'TRUE'
    };
  });

  return construirResposta(tarefas);
}

// Salva uma nova tarefa (método POST)
function doPost(e) {
  try {
    var corpo = JSON.parse(e.postData.contents);

    var id         = String(Date.now());
    var tarefa     = corpo.tarefa     || '';
    var descricao  = corpo.descricao  || '';
    var prioridade = corpo.prioridade || 'Média';
    var subtarefas = JSON.stringify(corpo.subtarefas || []);
    var concluido  = false;

    var planilha = getAba();
    planilha.appendRow([id, tarefa, descricao, prioridade, subtarefas, concluido]);

    return construirResposta({ sucesso: true, id: id });
  } catch (erro) {
    return construirResposta({ erro: erro.message });
  }
}

// Retorna a aba "Tarefas", criando-a com cabeçalhos se não existir
function getAba() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName(NOME_ABA);

  if (!aba) {
    aba = ss.insertSheet(NOME_ABA);
    aba.appendRow(CABECALHOS);
  }

  return aba;
}

// Tenta parsear JSON; retorna array vazio se falhar
function parsearJSON(valor) {
  try {
    return JSON.parse(valor);
  } catch (_) {
    return [];
  }
}

// Monta a resposta JSON
function construirResposta(dados) {
  return ContentService
    .createTextOutput(JSON.stringify(dados))
    .setMimeType(ContentService.MimeType.JSON);
}
