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

var NOME_ABA         = 'Tarefas';
var NOME_ABA_MEMBROS = 'Membros';
var CABECALHOS         = ['ID', 'Tarefa', 'Descrição', 'Prioridade', 'Subtarefas', 'Concluído', 'Vencimento', 'Responsável'];
var CABECALHOS_MEMBROS = ['ID', 'Nome'];

// Roteador GET: ?acao=membros → lista membros; padrão → lista tarefas
function doGet(e) {
  var acao = e && e.parameter && e.parameter.acao ? e.parameter.acao : 'tarefas';

  if (acao === 'membros') return listarMembros();
  return listarTarefas();
}

function listarTarefas() {
  var planilha = getAba();
  var dados    = planilha.getDataRange().getValues();

  var tarefas = dados.slice(1).map(function(linha) {
    return {
      id:          String(linha[0]),
      tarefa:      linha[1],
      descricao:   linha[2],
      prioridade:  linha[3],
      subtarefas:  parsearJSON(linha[4]),
      concluido:   linha[5] === true || linha[5] === 'TRUE',
      vencimento:  linha[6] ? String(linha[6]) : '',
      responsavel: linha[7] ? String(linha[7]) : ''
    };
  });

  return construirResposta(tarefas);
}

function listarMembros() {
  var aba   = getAbaMembros();
  var dados = aba.getDataRange().getValues();

  var membros = dados.slice(1).map(function(linha) {
    return { id: String(linha[0]), nome: linha[1] };
  });

  return construirResposta(membros);
}

// Roteador de ações POST
function doPost(e) {
  try {
    var corpo = JSON.parse(e.postData.contents);
    var acao  = corpo.acao || 'criar';

    if (acao === 'criar')          return acaoCriar(corpo);
    if (acao === 'concluir')       return acaoConcluir(corpo);
    if (acao === 'excluir')        return acaoExcluir(corpo);
    if (acao === 'adicionarMembro') return acaoAdicionarMembro(corpo);
    if (acao === 'excluirMembro')   return acaoExcluirMembro(corpo);

    return construirResposta({ erro: 'Ação desconhecida: ' + acao });
  } catch (erro) {
    return construirResposta({ erro: erro.message });
  }
}

// Cria uma nova tarefa
function acaoCriar(corpo) {
  var id         = String(Date.now());
  var tarefa     = corpo.tarefa     || '';
  var descricao  = corpo.descricao  || '';
  var prioridade = corpo.prioridade || 'Média';
  var subtarefas = JSON.stringify(corpo.subtarefas || []);
  var concluido  = false;
  var vencimento  = corpo.vencimento  || '';
  var responsavel = corpo.responsavel || '';

  var planilha = getAba();
  planilha.appendRow([id, tarefa, descricao, prioridade, subtarefas, concluido, vencimento, responsavel]);

  return construirResposta({ sucesso: true, id: id });
}

// Atualiza o campo Concluído de uma tarefa existente
function acaoConcluir(corpo) {
  var id        = String(corpo.id);
  var concluido = corpo.concluido === true;

  var planilha = getAba();
  var dados    = planilha.getDataRange().getValues();

  for (var i = 1; i < dados.length; i++) {
    if (String(dados[i][0]) === id) {
      // Coluna 6 = índice 5 (base 0), mas getRange usa base 1 → linha i+1, coluna 6
      planilha.getRange(i + 1, 6).setValue(concluido);
      return construirResposta({ sucesso: true });
    }
  }

  return construirResposta({ erro: 'Tarefa não encontrada: ' + id });
}

// Adiciona um membro da família
function acaoAdicionarMembro(corpo) {
  var nome = (corpo.nome || '').trim();
  if (!nome) return construirResposta({ erro: 'Nome obrigatório' });

  var aba   = getAbaMembros();
  var dados = aba.getDataRange().getValues();

  // Impede nome duplicado (case-insensitive)
  for (var i = 1; i < dados.length; i++) {
    if (String(dados[i][1]).toLowerCase() === nome.toLowerCase()) {
      return construirResposta({ erro: 'Membro já existe' });
    }
  }

  var id = String(Date.now());
  aba.appendRow([id, nome]);
  return construirResposta({ sucesso: true, id: id, nome: nome });
}

// Remove um membro pelo ID
function acaoExcluirMembro(corpo) {
  var id    = String(corpo.id);
  var aba   = getAbaMembros();
  var dados = aba.getDataRange().getValues();

  for (var i = 1; i < dados.length; i++) {
    if (String(dados[i][0]) === id) {
      aba.deleteRow(i + 1);
      return construirResposta({ sucesso: true });
    }
  }

  return construirResposta({ erro: 'Membro não encontrado' });
}

// Retorna a aba "Membros", criando-a se não existir
function getAbaMembros() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName(NOME_ABA_MEMBROS);

  if (!aba) {
    aba = ss.insertSheet(NOME_ABA_MEMBROS);
    aba.appendRow(CABECALHOS_MEMBROS);
  }

  return aba;
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

// Remove uma tarefa pelo ID
function acaoExcluir(corpo) {
  var id       = String(corpo.id);
  var planilha = getAba();
  var dados    = planilha.getDataRange().getValues();

  for (var i = 1; i < dados.length; i++) {
    if (String(dados[i][0]) === id) {
      planilha.deleteRow(i + 1);
      return construirResposta({ sucesso: true });
    }
  }

  return construirResposta({ erro: 'Tarefa não encontrada: ' + id });
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
