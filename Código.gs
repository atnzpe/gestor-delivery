// -------------------------------
// 1. CONSTANTES GLOBAIS
// -------------------------------
const ID_PLANILHA_MESTRA = SpreadsheetApp.getActiveSpreadsheet().getId();

const ABA_ESTOQUE = 'Estoque';
const ABA_CARDAPIO = 'Cardapio';
const ABA_VENDAS = 'Vendas';
const ABA_ADICIONAIS = 'Adicionais';
const ABA_PRODUTO_ADICIONAIS_LINK = 'Produto_Adicionais_Link';
const ABA_TAXA_ENTREGA = 'Taxa_Entrega';
const ABA_FICHA_TECNICA = 'Ficha Tecnica';
const ABA_CONFIG = 'Config';
const ABA_COMPRAS = 'Compras';
const ABA_LOG_ESTOQUE = 'LogEstoque';
const ABA_PAGAMENTOS = 'Pagamentos';
const ABA_CUSTO_INSUMOS = 'Custo_Insumos';
const ABA_PRECO_COMPRA_UNIDADE = 'Preço_CompraxUnidade';

// ====================================================================
// 2. ACESSO A DADOS (CORE)
// ====================================================================

function getPlanilha() {
  try {
    return SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {
    Logger.log(`[CORE][CRITICAL] Erro ao acessar planilha: ${e.message}`);
    throw new Error('Erro fatal: Não foi possível acessar o banco de dados.');
  }
}

function getSheet(ss, name) {
  try {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      Logger.log(`[CORE][WARN] Aba "${name}" não encontrada. Tentando criar...`);
      sheet = ss.insertSheet(name);
    }
    return sheet;
  } catch (e) {
    Logger.log(`[CORE][ERROR] Falha ao acessar aba ${name}: ${e.message}`);
    throw new Error(`Erro interno ao acessar tabela: ${name}`);
  }
}

function getSheetOrCreate(ss, name, headers) {
  try {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      if (headers && headers.length > 0) sheet.appendRow(headers);
      Logger.log(`[CORE][INFO] Aba criada: ${name}`);
    }
    return sheet;
  } catch (e) {
    throw new Error(`Erro ao estruturar tabela ${name}: ${e.message}`);
  }
}

// ====================================================================
// 3. CONFIGURAÇÃO E SETUP (AUTO-CRIAÇÃO)
// ====================================================================

/**
 * @function configurarPlanilha
 * @description Cria abas e cabeçalhos automaticamente para o teste do zero.
 */
function configurarPlanilha() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const estrutura = [
    { nome: ABA_ESTOQUE, headers: ['ID_INSUMO', 'NOME', 'UNIDADE', 'ESTOQUE_ATUAL', 'ATIVO'] },
    { nome: ABA_CARDAPIO, headers: ['ID_ITEM', 'NOME', 'PRECO', 'CATEGORIA', 'CUSTO', 'ATIVO', 'TEMPO_PREPARO', 'DISPONIVEL_DELIVERY', 'PRECO_SUGERIDO', 'MARGEM_LUCRO', 'CUSTO_PERCENTUAL', 'LUCRO_BRUTO', 'DESCRICAO', 'FOTO_URL'] },
    { nome: ABA_VENDAS, headers: ['ID_VENDA', 'DT_HORA_PEDIDO', 'JSON_PEDIDO', 'VL_TOTAL_PEDIDO', 'FORMAPAGAMENTO_PEDIDO', 'CLIENTE_NOME', 'CLIENTE_TELEFONE', 'LOGRADOURO', 'NUMERO', 'COMPLEMENTO', 'BAIRRO', 'CIDADE', 'PONTO_REFERENCIA', 'TAXA_ENTREGA', 'STATUS_PEDIDO', 'OBSERVACOES'] },
    { nome: ABA_ADICIONAIS, headers: ['ID_ADICIONAL', 'NOME', 'PRECO', 'CATEGORIA', 'ATIVO'] },
    { nome: ABA_PRODUTO_ADICIONAIS_LINK, headers: ['ID_PRODUTO', 'ID_ADICIONAL', 'ATIVO'] },
    { nome: ABA_FICHA_TECNICA, headers: ['ID_FICHA', 'ID_ITEM_CARDAPIO_FK', 'ID_INSUMO_FK', 'QUANTIDADE_USADA', 'NOME_ITEM_CACHE', 'NOME_INSUMO_CACHE'] },
    { nome: ABA_CONFIG, headers: ['CHAVE', 'VALOR'] },
    { nome: ABA_COMPRAS, headers: ['DT_HORA_COMPRA', 'ID_INSUMO', 'QTD_COMPRA', 'PRECO_TOTAL', 'NOME_INSUMO'] },
    { nome: ABA_LOG_ESTOQUE, headers: ['DT_HORA_LOG', 'ID_INSUMO', 'QUANTIDADE', 'MOTIVO', 'NOME_INSUMO'] },
    { nome: ABA_TAXA_ENTREGA, headers: ['BAIRRO', 'TAXA'] },
    { nome: ABA_CUSTO_INSUMOS, headers: ['NOME_INSUMO', 'ID_INSUMO', 'QTD_ULTIMA_COMPRA', 'PRECO_ULTIMA_COMPRA', 'DATA_ATUALIZACAO', 'CUSTO_UNITARIO'] },
    { nome: ABA_PRECO_COMPRA_UNIDADE, headers: ['ITEM', 'UNIDADES_DE_COMPRA', 'QTD_COMPRADA', 'PRECO_COMPRA', 'PRECO_TOTAL', 'PRECO_UNITARIO'] },
    { nome: ABA_PAGAMENTOS, headers: ['METODO', 'ATIVO'] },
  ];

  estrutura.forEach(obj => {
    let sheet = ss.getSheetByName(obj.nome);
    if (!sheet) {
      sheet = ss.insertSheet(obj.nome);
      sheet.appendRow(obj.headers);
      Logger.log(`[SETUP] Aba '${obj.nome}' criada com sucesso.`);
    } else {
      if (sheet.getLastRow() === 0) {
        sheet.appendRow(obj.headers);
      }
    }
  });

  const abaConfig = getSheetOrCreate(ss, ABA_CONFIG);
  if (abaConfig.getLastRow() <= 1) {
    abaConfig.appendRow(['NumeroWhatsApp', '5511999999999']);
  }

  SpreadsheetApp.getUi().alert('✅ Ambiente POC Configurado! Todas as abas e cabeçalhos foram recriados.');
}

// ====================================================================
// 4. MENU E UI
// ====================================================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🍴 Gerenciar Restaurante')
    .addSubMenu(ui.createMenu('Cadastros (Criar)')
      .addItem('Insumo (Estoque)', 'showSidebarCadastroInsumo')
      .addItem('Item Cardápio', 'showSidebarCadastroCardapio')
      .addItem('Adicional', 'showSidebarGerenciarAdicionais')
      .addItem('Ligar Adicionais', 'showSidebarLigarAdicionais')
      .addItem('Ficha Técnica', 'showSidebarCadastroFichaTecnica')
      .addSeparator()
      .addItem('Taxas de Entrega', 'showSidebarGerenciarTaxas')
      .addItem('Formas de Pagamento', 'showSidebarGerenciarPagamentos'))
    .addSeparator()
    .addSubMenu(ui.createMenu('Gerenciar (Editar)')
      .addItem('Cardápio', 'showSidebarGerenciarCardapio')
      .addItem('Insumos', 'showSidebarGerenciarInsumos')
      .addItem('Adicionais', 'showSidebarGerenciarAdicionais'))
    .addSeparator()
    .addSubMenu(ui.createMenu('Estoque & Custos')
      .addItem('Entrada de Nota (Lote)', 'showSidebarRegistrarCompraLote')
      .addItem('Ajuste Manual', 'showSidebarAjustarEstoque')
      .addItem('Recalcular Custos', 'uiRecalcularCustoDosProdutos')
      .addItem('📊 Atualizar Dashboard', 'uiAtualizarDashboard')
      .addItem('📑 Central de Relatórios', 'showSidebarRelatorios'))
    .addSeparator()
    .addItem('📱 Configurar WhatsApp', 'showSidebarConfigWhatsApp')
    .addItem('🚨 LIMPAR DADOS TESTE', 'adminLimparDadosDeTeste')
    .addToUi();
}

// Wrappers
function showSidebarCadastroInsumo() { openSidebar('CadastroInsumo', 'Novo Insumo'); }
function showSidebarCadastroCardapio() { openSidebar('CadastroCardapio', 'Novo Item Cardápio'); }
function showSidebarCadastroFichaTecnica() { openSidebar('CadastroFichaTecnica', 'Ficha Técnica'); }
function showSidebarGerenciarAdicionais() { openSidebar('GerenciarAdicionais', 'Gerenciar Adicionais'); }
function showSidebarLigarAdicionais() { openSidebar('LigarAdicionais', 'Vincular Adicionais'); }
function showSidebarGerenciarTaxas() { openSidebar('GerenciarTaxas', 'Taxas de Entrega'); }
function showSidebarGerenciarPagamentos() { openSidebar('GerenciarPagamentos', 'Pagamentos'); }
function showSidebarConfigWhatsApp() { openSidebar('ConfigWhatsApp', 'Configurar WhatsApp'); }
function showSidebarGerenciarCardapio() { openSidebar('GerenciarCardapio', 'Gerenciar Cardápio'); }
function showSidebarGerenciarInsumos() { openSidebar('GerenciarInsumos', 'Gerenciar Insumos'); }
function showSidebarRegistrarCompraLote() { openSidebar('RegistrarCompraLote', 'Entrada de Nota'); }
function showSidebarAjustarEstoque() { openSidebar('AjustarEstoque', 'Ajuste de Estoque'); }
function showSidebarRelatorios() { openSidebar('CentralRelatorios', 'Central de Relatórios'); }

function openSidebar(template, title) {
  const html = getHtmlTemplate(template).setTitle(title);
  SpreadsheetApp.getUi().showSidebar(html);
}

function uiRecalcularCustoDosProdutos() {
  recalcularCustoDosProdutos(null);
  SpreadsheetApp.getUi().alert('Custos recalculados!');
}

// ====================================================================
// MÓDULO: GESTÃO DE INSUMOS (Refatorado com QA & Logs)
// ====================================================================

/**
 * Processa o cadastro de um novo insumo com validações rigorosas.
 * @param {Object} formData - Dados vindos do formulário HTML
 */
function processarCadastroInsumo(formData) {
  // [DEBUG] Rastreabilidade de entrada
  Logger.log(`[INSUMO][CREATE] Iniciando cadastro. Payload recebido: ${JSON.stringify(formData)}`);

  try {
    // 1. Validação de Campos (QA - Prevenção de Erros)
    if (!formData.nome || formData.nome.trim() === "") {
      throw new Error("O nome do insumo é obrigatório.");
    }
    if (!formData.unidade || formData.unidade.trim() === "") {
      throw new Error("A unidade de medida é obrigatória.");
    }

    const ss = getPlanilha();
    const aba = getSheet(ss, ABA_ESTOQUE);
    const lastRow = aba.getLastRow();

    // 2. Geração de ID (Pode ser melhorado com Timestamp para evitar colisão)
    const novoId = `IN-${lastRow}`;

    // 3. Sanitização de Dados (DCU - Aceitar vírgula ou ponto)
    // Garante que o usuário não precise se preocupar com o formato numérico do teclado
    let estoqueStr = String(formData.estoque || '0').replace(',', '.');
    let estoqueInicial = parseFloat(estoqueStr);

    // Proteção contra NaN (Not a Number)
    if (isNaN(estoqueInicial)) {
      Logger.log(`[INSUMO][WARN] Valor de estoque inválido: "${formData.estoque}". Forçando 0.`);
      estoqueInicial = 0;
    }

    // 4. Persistência
    aba.appendRow([
      novoId,
      String(formData.nome).trim(),
      String(formData.unidade).trim().toUpperCase(),
      estoqueInicial,
      'SIM'
    ]);

    // 5. Log de Estoque (Rastreabilidade)
    if (estoqueInicial > 0) {
      registrarLogEstoque(novoId, formData.nome, estoqueInicial, 'Cadastro Inicial');
    } else {
      // Registra mesmo zerado para confirmar a criação do item no histórico
      registrarLogEstoque(novoId, formData.nome, 0, 'Item Criado (Estoque Zero)');
    }

    // [DEBUG] Sucesso
    Logger.log(`[INSUMO][CREATE] Sucesso. ID Gerado: ${novoId}`);
    return 'Insumo cadastrado com sucesso!';

  } catch (e) {
    // [DEBUG] Erro
    Logger.log(`[INSUMO][ERROR] Falha ao cadastrar: ${e.message} | Stack: ${e.stack}`);
    // Repassa erro amigável para o frontend
    throw new Error(`Erro ao salvar insumo: ${e.message}`);
  }
}

/**
 * Busca insumos para listagem no admin com tratamento de erros.
 */
function getInsumosParaAdmin() {
  Logger.log(`[INSUMO][READ] Buscando lista de insumos...`);

  try {
    const aba = getSheet(getPlanilha(), ABA_ESTOQUE);

    // Verifica se há dados além do cabeçalho
    if (aba.getLastRow() <= 1) {
      Logger.log(`[INSUMO][READ] Tabela vazia.`);
      return [];
    }

    const data = aba.getDataRange().getValues();
    data.shift(); // Remove cabeçalho

    const listaFormatada = data.map((r, i) => ({
      rowIndex: i + 2, // Mantém referência da linha (Base 1 + Header)
      id: r[0],
      nome: r[1],
      unidade: r[2],
      estoqueAtual: Number(r[3]) || 0, // Garante número
      ativo: r[4]
    })).filter(x => x.id); // Remove linhas vazias/fantasmas

    Logger.log(`[INSUMO][READ] Retornados ${listaFormatada.length} insumos.`);
    return listaFormatada;

  } catch (e) {
    Logger.log(`[INSUMO][ERROR] Falha ao listar: ${e.message}`);
    throw new Error("Não foi possível carregar a lista de insumos.");
  }
}

/**
 * Atualiza um insumo existente com verificação de integridade.
 */
function updateInsumo(formData) {
  Logger.log(`[INSUMO][UPDATE] Iniciando atualização. Dados: ${JSON.stringify(formData)}`);

  try {
    const sheet = getSheet(getPlanilha(), ABA_ESTOQUE);
    const row = Number(formData.rowIndex);

    // 1. Validação de Integridade (QA Crítico)
    // Verifica se a linha que vamos editar ainda contém o mesmo ID.
    // Isso evita editar o item errado se alguém tiver ordenado a planilha nesse meio tempo.
    const idNaPlanilha = sheet.getRange(row, 1).getValue(); // Coluna A é o ID

    if (String(idNaPlanilha) !== String(formData.id)) {
      Logger.log(`[INSUMO][CRITICAL] Conflito de IDs. Esperado: ${formData.id}, Encontrado: ${idNaPlanilha}`);
      throw new Error("A planilha foi modificada externamente. Por favor, recarregue a página e tente novamente.");
    }

    // 2. Persistência
    // Coluna 2 = Nome, 3 = Unidade, 5 = Ativo
    sheet.getRange(row, 2).setValue(formData.nome);
    sheet.getRange(row, 3).setValue(formData.unidade);
    sheet.getRange(row, 5).setValue(formData.ativo);

    Logger.log(`[INSUMO][UPDATE] Sucesso. Linha ${row} atualizada.`);
    return 'Insumo atualizado com sucesso!';

  } catch (e) {
    Logger.log(`[INSUMO][ERROR] Falha ao atualizar: ${e.message}`);
    throw new Error(`Erro ao atualizar: ${e.message}`);
  }
}

/// ====================================================================
// MÓDULO: GESTÃO DE CARDÁPIO (Refatorado com QA & Logs)
// ====================================================================

/**
 * Cadastra um novo produto no cardápio com validação de dados.
 */
function processarCadastroCardapio(data) {
  Logger.log(`[CARDAPIO][CREATE] Iniciando cadastro. Payload: ${JSON.stringify(data)}`);

  try {
    // 1. Validação de Entrada (QA)
    if (!data.nome || data.nome.trim().length < 2) {
      throw new Error("O nome do produto é muito curto ou vazio.");
    }
    if (!data.categoria || data.categoria.trim() === "") {
      throw new Error("Selecione ou digite uma categoria.");
    }

    // 2. Sanitização de Preço (DCU)
    // Aceita "10,50" ou "10.50"
    let precoStr = String(data.preco).replace(',', '.');
    let precoFinal = parseFloat(precoStr);

    if (isNaN(precoFinal) || precoFinal < 0) {
      throw new Error("Preço inválido. Digite um valor numérico.");
    }

    const ss = getPlanilha();
    const sheet = getSheet(ss, ABA_CARDAPIO);

    // 3. Geração de ID
    // Usa lastRow + Timestamp curto para evitar duplicidade se apagar linhas do meio
    const lastRow = sheet.getLastRow();
    const novoId = `CD-${lastRow + 1}`;

    // 4. Persistência
    sheet.appendRow([
      novoId,
      String(data.nome).trim(),
      precoFinal,
      String(data.categoria).trim(),
      0, // Custo (Calculado depois)
      String(data.ativo || 'SIM').toUpperCase(),
      '', '', '', '', '', '', // Colunas G a L vazias (reservadas)
      String(data.descricao || ''),
      String(data.foto_url || '')
    ]);

    Logger.log(`[CARDAPIO][CREATE] Sucesso. ID: ${novoId} - Item: ${data.nome}`);
    return 'Produto salvo com sucesso!';

  } catch (e) {
    Logger.log(`[CARDAPIO][ERROR] ${e.message}`);
    throw new Error(`Erro ao cadastrar: ${e.message}`);
  }
}

/**
 * Busca produtos para a tabela do Admin com tratamento de erro.
 */
function getCardapioParaAdmin() {
  Logger.log(`[CARDAPIO][READ] Buscando lista...`);
  try {
    const sheet = getSheet(getPlanilha(), ABA_CARDAPIO);

    if (sheet.getLastRow() <= 1) {
      Logger.log(`[CARDAPIO][READ] Cardápio vazio.`);
      return [];
    }

    const data = sheet.getDataRange().getValues();
    data.shift(); // Remove Header

    const lista = data.map((r, i) => ({
      rowIndex: i + 2, // Referência absoluta da linha
      id: r[0],
      nome: r[1],
      preco: Number(r[2]) || 0,
      categoria: r[3],
      custo: Number(r[4]) || 0,
      ativo: r[5],
      descricao: r[12], // Coluna M
      foto_url: r[13]   // Coluna N
    })).filter(x => x.id); // Remove linhas em branco

    return lista;

  } catch (e) {
    Logger.log(`[CARDAPIO][ERROR] Falha na listagem: ${e.message}`);
    throw new Error("Não foi possível carregar o cardápio.");
  }
}

/**
 * Atualiza um produto existente com Trava de Segurança de ID.
 */
function updateProdutoCardapio(data) {
  Logger.log(`[CARDAPIO][UPDATE] Atualizando ID: ${data.id}`);

  try {
    const sheet = getSheet(getPlanilha(), ABA_CARDAPIO);
    const row = Number(data.rowIndex);

    // 1. Validação de Integridade (QA Crítico)
    // Verifica se a linha ainda pertence ao mesmo produto.
    // Impede que uma reordenação da planilha faça você editar o item errado.
    const idNaPlanilha = sheet.getRange(row, 1).getValue(); // Lê coluna A da linha alvo

    if (String(idNaPlanilha) !== String(data.id)) {
      Logger.log(`[CARDAPIO][CRITICAL] Conflito de IDs. Planilha: ${idNaPlanilha} != Form: ${data.id}`);
      throw new Error("A planilha mudou de posição. Recarregue a página para editar com segurança.");
    }

    // 2. Sanitização
    let precoFinal = parseFloat(String(data.preco).replace(',', '.'));
    if (isNaN(precoFinal)) precoFinal = 0;

    // 3. Atualização (Mapeamento explícito de colunas)
    sheet.getRange(row, 2).setValue(data.nome);       // Col B
    sheet.getRange(row, 3).setValue(precoFinal);      // Col C
    sheet.getRange(row, 4).setValue(data.categoria);  // Col D
    sheet.getRange(row, 6).setValue(data.ativo);      // Col F
    sheet.getRange(row, 13).setValue(data.descricao); // Col M
    sheet.getRange(row, 14).setValue(data.foto_url);  // Col N

    Logger.log(`[CARDAPIO][UPDATE] Sucesso na linha ${row}.`);
    return 'Produto atualizado com sucesso!';

  } catch (e) {
    Logger.log(`[CARDAPIO][ERROR] ${e.message}`);
    throw new Error(`Erro ao atualizar: ${e.message}`);
  }
}

// ====================================================================
// MÓDULO: GESTÃO DE ADICIONAIS & VÍNCULOS (Refatorado)
// ====================================================================

/**
 * Cria um novo adicional (Ex: Bacon, Cheddar).
 */
function criarAdicional(formData) {
  Logger.log(`[ADICIONAL][CREATE] Iniciando cadastro: ${formData.nome}`);
  const lock = LockService.getScriptLock();

  try {
    // Tenta obter lock por 5s para evitar duplicidade simultânea
    if (!lock.tryLock(5000)) {
      throw new Error("Sistema ocupado. Tente novamente.");
    }

    // Validação Básica
    if (!formData.nome) throw new Error("Nome é obrigatório.");
    let preco = parseFloat(String(formData.preco).replace(',', '.'));
    if (isNaN(preco)) preco = 0;

    const sheet = getSheetOrCreate(getPlanilha(), ABA_ADICIONAIS, ["ID_ADICIONAL", "NOME", "PRECO", "CATEGORIA", "ATIVO"]);
    const lastRow = sheet.getLastRow();
    const id = `AD-${lastRow + 1}`; // ID Sequencial

    sheet.appendRow([id, formData.nome.trim(), preco, formData.categoria || 'Geral', 'SIM']);

    Logger.log(`[ADICIONAL][CREATE] Sucesso. ID: ${id}`);
    return 'Adicional criado com sucesso!';

  } catch (e) {
    Logger.log(`[ADICIONAL][ERROR] ${e.message}`);
    throw new Error(`Erro ao criar adicional: ${e.message}`);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Lista adicionais para a tabela administrativa.
 */
function getAdicionaisParaAdmin() {
  try {
    const sheet = getSheet(getPlanilha(), ABA_ADICIONAIS);
    if (sheet.getLastRow() <= 1) return [];

    const data = sheet.getDataRange().getValues();
    data.shift(); // Remove header

    return data.map((r, i) => ({
      rowIndex: i + 2,
      id: r[0],
      nome: r[1],
      preco: Number(r[2]) || 0,
      categoria: r[3],
      ativo: r[4]
    })).filter(x => x.id);

  } catch (e) {
    Logger.log(`[ADICIONAL][ERROR] Falha na listagem: ${e.message}`);
    throw new Error("Erro ao carregar lista.");
  }
}

/**
 * Atualiza um adicional existente com verificação de ID.
 */
function updateAdicional(formData) {
  Logger.log(`[ADICIONAL][UPDATE] Editando ID: ${formData.id}`);

  try {
    const sheet = getSheet(getPlanilha(), ABA_ADICIONAIS);
    const row = Number(formData.rowIndex);

    // Validação de Integridade (QA)
    // Garante que a linha ainda pertence ao mesmo ID antes de escrever
    const currentId = sheet.getRange(row, 1).getValue();
    if (String(currentId) !== String(formData.id)) {
      // Se o ID não bate, tenta buscar a linha correta (Fallback de segurança)
      Logger.log(`[ADICIONAL][WARN] ID na linha ${row} não confere. Buscando novamente...`);
      const data = sheet.getDataRange().getValues();
      let foundRow = -1;
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(formData.id)) {
          foundRow = i + 1;
          break;
        }
      }
      if (foundRow === -1) throw new Error("Adicional não encontrado ou excluído.");

      // Redireciona para a linha correta encontrada
      sheet.getRange(foundRow, 2).setValue(formData.nome);
      sheet.getRange(foundRow, 3).setValue(parseFloat(String(formData.preco).replace(',', '.')));
      sheet.getRange(foundRow, 4).setValue(formData.categoria);
      sheet.getRange(foundRow, 5).setValue(formData.ativo);
      return 'Adicional atualizado (Linha corrigida)!';
    }

    // Caminho feliz (Linha correta)
    sheet.getRange(row, 2).setValue(formData.nome);
    sheet.getRange(row, 3).setValue(parseFloat(String(formData.preco).replace(',', '.')));
    sheet.getRange(row, 4).setValue(formData.categoria);
    sheet.getRange(row, 5).setValue(formData.ativo);

    return 'Adicional atualizado!';

  } catch (e) {
    Logger.log(`[ADICIONAL][ERROR] ${e.message}`);
    throw new Error(e.message);
  }
}

/**
 * Helper para soft-delete.
 */
function inativarAdicional(formData) {
  formData.ativo = 'NAO';
  return updateAdicional(formData);
}

// --- VÍNCULOS (PRODUTO <-> ADICIONAL) ---

function getDadosParaLigarAdicionais() {
  // Reutiliza as funções blindadas de leitura
  const prods = getCardapioParaAdmin().filter(p => p.ativo === 'SIM');
  const adds = getAdicionaisParaAdmin().filter(a => a.ativo === 'SIM');
  return { produtos: prods, adicionais: adds };
}

function getLinksAtuais(idProduto) {
  try {
    const sheet = getSheetOrCreate(getPlanilha(), ABA_PRODUTO_ADICIONAIS_LINK, ['ID_PRODUTO', 'ID_ADICIONAL', 'ATIVO']);
    const data = sheet.getDataRange().getValues();
    // Filtra apenas os ativos ('SIM') para este produto
    return data
      .filter(r => String(r[0]) === String(idProduto) && r[2] === 'SIM')
      .map(r => r[1]);
  } catch (e) {
    return [];
  }
}

/**
 * Salva os vínculos (N para N) de forma otimizada.
 */
function processarLinkAdicionais(data) {
  Logger.log(`[LINK][UPDATE] Atualizando vínculos para Produto: ${data.idProduto}`);
  const lock = LockService.getScriptLock();

  try {
    if (!lock.tryLock(5000)) throw new Error("Sistema ocupado.");

    const sheet = getSheetOrCreate(getPlanilha(), ABA_PRODUTO_ADICIONAIS_LINK, ['ID_PRODUTO', 'ID_ADICIONAL', 'ATIVO']);
    const idProd = data.idProduto;
    const novosAdds = data.adicionais || []; // Array de IDs

    // 1. Lê tudo
    const allData = sheet.getDataRange().getValues();
    const header = allData.shift(); // Remove header

    // 2. Filtra mantendo apenas o que NÃO É deste produto
    // (Estratégia: Apagar tudo deste produto e recriar apenas os selecionados)
    const outrosLinks = allData.filter(r => String(r[0]) !== String(idProd));

    // 3. Adiciona os novos selecionados
    const novasLinhas = novosAdds.map(idAdd => [idProd, idAdd, 'SIM']);

    // 4. Reconstrói a tabela
    const finalData = [header, ...outrosLinks, ...novasLinhas];

    // 5. Gravação Atômica (Limpa e Cola)
    sheet.clearContents();
    if (finalData.length > 0) {
      sheet.getRange(1, 1, finalData.length, 3).setValues(finalData);
    }

    Logger.log(`[LINK][SUCCESS] ${novasLinhas.length} adicionais vinculados.`);
    return 'Vínculos salvos com sucesso!';

  } catch (e) {
    Logger.log(`[LINK][ERROR] ${e.message}`);
    throw new Error(`Erro ao salvar vínculos: ${e.message}`);
  } finally {
    lock.releaseLock();
  }
}

// === FIM ====



function processarRegistroCompra(formData) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = getPlanilha();
    const abaCompras = getSheet(ss, ABA_COMPRAS);
    const abaEstoque = getSheet(ss, ABA_ESTOQUE);
    const abaCusto = getSheet(ss, ABA_CUSTO_INSUMOS); // Aba pode não existir no POC original, mas foi criada no setup

    const idInsumo = formData.idInsumo;
    const qtde = Number(formData.quantidade);
    const preco = Number(formData.precoCompra);
    const total = qtde * preco;
    const nome = getNomeInsumoPorID(abaEstoque, idInsumo);

    abaCompras.appendRow([new Date(), idInsumo, qtde, total, nome]);
    atualizarEstoque(abaEstoque, idInsumo, qtde);
    registrarLogEstoque(idInsumo, nome, qtde, 'Compra');

    // Lógica simples de atualização de custo (Último Preço)
    // Em produção, usar Média Ponderada
    // Aqui vamos apenas salvar na aba de Custos se ela existir
    if (abaCusto) {
      // Lógica simplificada: adiciona linha de custo
      abaCusto.appendRow([nome, idInsumo, qtde, preco, new Date(), preco]);
    }

    SpreadsheetApp.flush();
    recalcularCustoDosProdutos(null);
    return 'Compra registrada e custos atualizados!';
  } finally {
    lock.releaseLock();
  }
}


function gerarRelatorio(tipo) {
  const ss = getPlanilha();
  const dataHora = new Date().toLocaleString();
  let nomeAba = tipo === 'ESTOQUE_BAIXO' ? 'REL_ESTOQUE' : 'REL_VENDAS';
  let dadosRelatorio = [];
  let headers = [];

  if (tipo === 'ESTOQUE_BAIXO') {
    headers = ['ID', 'Insumo', 'Unid', 'Estoque Atual', 'Status'];
    const estoque = getSheet(ss, ABA_ESTOQUE).getDataRange().getValues();
    estoque.shift();
    dadosRelatorio = estoque.filter(r => r[4] === 'SIM').map(r => {
      const qtd = parseFloat(r[3]);
      let status = qtd <= 0 ? 'CRÍTICO' : (qtd < 5 ? 'BAIXO' : 'OK');
      return [r[0], r[1], r[2], qtd, status];
    }).sort((a, b) => a[3] - b[3]);
  } else if (tipo === 'VENDAS_DETALHADA') {
    headers = ['Data', 'Cliente', 'Bairro', 'Pagamento', 'Total', 'Status'];
    const vendas = getSheet(ss, ABA_VENDAS).getDataRange().getValues();
    vendas.shift();
    dadosRelatorio = vendas.reverse().slice(0, 100).map(r => [new Date(r[1]).toLocaleString(), r[5], r[10], r[4], parseFloat(r[3]), r[14]]);
  }

  const oldSheet = ss.getSheetByName(nomeAba);
  if (oldSheet) ss.deleteSheet(oldSheet);
  const sheet = ss.insertSheet(nomeAba);

  sheet.getRange("A1").setValue(`Relatório: ${tipo}`);
  if (dadosRelatorio.length > 0) {
    sheet.getRange(3, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.getRange(4, 1, dadosRelatorio.length, headers.length).setValues(dadosRelatorio);
  } else {
    sheet.getRange("A3").setValue("Sem dados.");
  }
  return `Relatório gerado em ${nomeAba}`;
}

// ====================================================================
// 11. MÓDULO COMPRA EM LOTE (EFICIÊNCIA)
// ====================================================================

function showSidebarRegistrarCompraLote() {
  const html = getHtmlTemplate('RegistrarCompraLote');
  SpreadsheetApp.getUi().showSidebar(html.setTitle('Entrada de Nota Fiscal').setWidth(450));
}

function processarCompraLote(itensCompra) {
  const lock = LockService.getScriptLock();
  // Tenta segurar o sistema por 30s para garantir integridade
  if (!lock.tryLock(30000)) {
    throw new Error('Sistema ocupado. Tente novamente em alguns segundos.');
  }

  try {
    const ss = getPlanilha();
    const abaCompras = getSheet(ss, ABA_COMPRAS);
    const abaEstoque = getSheet(ss, ABA_ESTOQUE);
    const abaCusto = getSheet(ss, ABA_CUSTO_INSUMOS);

    const dataHora = new Date();
    const dadosEstoque = abaEstoque.getDataRange().getValues();
    const dadosCusto = abaCusto.getDataRange().getValues(); // Cache para performance

    // Array para salvar em lote na aba Compras
    const novasLinhasCompra = [];

    itensCompra.forEach(item => {
      // 1. Dados do item
      const idInsumo = item.idInsumo;
      const qtd = parseFloat(item.quantidade);
      const preco = parseFloat(item.preco);
      const total = qtd * preco;

      // 2. Busca Nome (Otimizado na memória)
      let nomeInsumo = 'Desconhecido';
      let linhaEstoqueIdx = -1;

      for (let i = 1; i < dadosEstoque.length; i++) {
        if (String(dadosEstoque[i][0]) === String(idInsumo)) {
          nomeInsumo = dadosEstoque[i][1];
          linhaEstoqueIdx = i + 1; // Base 1
          break;
        }
      }

      if (linhaEstoqueIdx === -1) return; // Pula se não achar ID

      // 3. Prepara linha de histórico
      novasLinhasCompra.push([dataHora, idInsumo, qtd, total, nomeInsumo]);

      // 4. Atualiza Estoque (Direto na célula para garantir)
      const celulaEstoque = abaEstoque.getRange(linhaEstoqueIdx, 4);
      const estoqueAtual = parseFloat(celulaEstoque.getValue()) || 0;
      celulaEstoque.setValue(estoqueAtual + qtd);

      // 5. Atualiza Custo (Direto na célula se achar)
      let linhaCustoIdx = -1;
      for (let k = 1; k < dadosCusto.length; k++) {
        if (dadosCusto[k][1] === idInsumo) { // Coluna B = ID
          linhaCustoIdx = k + 1;
          break;
        }
      }

      // Se achou custo, atualiza. Se não, cria.
      if (linhaCustoIdx > 0) {
        abaCusto.getRange(linhaCustoIdx, 3).setValue(qtd); // Qtd Ultima
        abaCusto.getRange(linhaCustoIdx, 4).setValue(preco); // Preço Ultima
        // Aqui poderia entrar lógica de média ponderada futura
        abaCusto.getRange(linhaCustoIdx, 6).setValue(preco); // Custo Unitário Atual
      } else {
        abaCusto.appendRow([nomeInsumo, idInsumo, qtd, preco, new Date(), preco]);
      }

      // 6. Log
      registrarLogEstoque(idInsumo, nomeInsumo, qtd, 'Entrada Nota (Lote)');
    });

    // Salva histórico em lote (Performance)
    if (novasLinhasCompra.length > 0) {
      abaCompras.getRange(abaCompras.getLastRow() + 1, 1, novasLinhasCompra.length, 5).setValues(novasLinhasCompra);
    }

    SpreadsheetApp.flush();
    recalcularCustoDosProdutos(null); // Atualiza preços do cardápio

    return `${itensCompra.length} itens processados com sucesso!`;

  } catch (e) {
    Logger.log(e);
    throw new Error(`Erro no lote: ${e.message}`);
  } finally {
    lock.releaseLock();
  }
}





function processarAjusteEstoque(formData) {
  const abaEstoque = getSheet(getPlanilha(), ABA_ESTOQUE);
  const atual = getEstoqueAtual(abaEstoque, formData.idInsumo);
  const real = Number(formData.contagemReal);
  const diff = real - atual;

  if (diff === 0) return 'Estoque já está correto.';

  const nome = getNomeInsumoPorID(abaEstoque, formData.idInsumo);
  setEstoque(abaEstoque, formData.idInsumo, real);
  registrarLogEstoque(formData.idInsumo, nome, diff, formData.motivo || 'Ajuste Manual');

  return `Estoque ajustado de ${atual} para ${real}.`;
}

// ====================================================================
// 6. GETTERS DE DADOS
// ====================================================================

function getDadosParaFormularios() {
  const ss = getPlanilha();
  // Insumos
  const iData = getSheet(ss, ABA_ESTOQUE).getDataRange().getValues();
  iData.shift();
  const insumos = iData
    .filter(r => r[0] && String(r[4]) === 'SIM')
    .map(r => ({ id: r[0], nome: r[1] }));

  // Cardapio
  const cData = getSheet(ss, ABA_CARDAPIO).getDataRange().getValues();
  cData.shift();
  const cardapio = cData
    .filter(r => r[0] && String(r[5]) === 'SIM')
    .map(r => ({ id: r[0], nome: r[1] }));

  return { insumos, cardapio };
}

// ====================================================================
// 🧩 MÓDULO: ENGENHARIA DE CARDÁPIO (FICHA TÉCNICA & CUSTOS)
// ====================================================================

/**
 * Vincula um insumo a um produto com quantidade definida.
 * @param {Object} data - { idItemCardapio, idInsumo, quantidade }
*/

function processarCadastroFichaTecnica(data) {
  
  Logger.log(`[FICHA][CREATE] Iniciando vínculo. Produto: ${data.idItemCardapio} + Insumo: ${data.idInsumo}`);

  try {
    // 1. Validação de Entrada (QA)
    if (!data.idItemCardapio || !data.idInsumo) {
      throw new Error("Produto e Insumo são obrigatórios.");
    }

    // Sanitização de número (DCU: aceita vírgula)
    const qtd = parseFloat(String(data.quantidade).replace(',', '.'));
    if (isNaN(qtd) || qtd <= 0) {
      throw new Error("A quantidade deve ser um número maior que zero.");
    }

    const sheet = getSheet(getPlanilha(), ABA_FICHA_TECNICA);
    const lastRow = sheet.getLastRow();
    const novoId = `FT-${lastRow + 1}`; // ID Sequencial

    // 2. Cache visual (Para leitura humana na planilha, opcional mas útil)
    const nomeItem = getNomeItemPorID(data.idItemCardapio) || 'Item Desconhecido';
    const nomeInsumo = getNomeInsumoPorID(getSheet(getPlanilha(), ABA_ESTOQUE), data.idInsumo) || 'Insumo Desconhecido';

    // 3. Persistência
    sheet.appendRow([
      novoId,
      data.idItemCardapio,
      data.idInsumo,
      qtd,
      nomeItem,   // Coluna E (Cache Nome Produto)
      nomeInsumo  // Coluna F (Cache Nome Insumo)
    ]);

    // 4. Recálculo Imediato (Regra de Negócio)
    // Atualiza o preço de custo do produto assim que o ingrediente entra
    recalcularCustoDosProdutos(data.idItemCardapio);

    Logger.log(`[FICHA][SUCCESS] Vínculo criado: ${novoId}`);
    return 'Ingrediente adicionado com sucesso!';

  } catch (e) {
    Logger.log(`[FICHA][ERROR] ${e.message}`);
    throw new Error(`Erro ao adicionar ingrediente: ${e.message}`);
  }
}

/**
 * Busca a lista de ingredientes de um produto para exibir no modal/tela.
 */
function getFichaTecnicaDetalhada(idProd) {
  Logger.log(`[FICHA][READ] Buscando receita do produto: ${idProd}`);

  try {
    const ss = getPlanilha();
    const sheet = getSheet(ss, ABA_FICHA_TECNICA);

    if (sheet.getLastRow() <= 1) return [];

    const dados = sheet.getDataRange().getValues();
    dados.shift(); // Remove cabeçalho

    // Otimização: Busca nomes dos insumos para não mostrar só IDs
    const insumosAdmin = getInsumosParaAdmin();
    const mapInsumos = {};
    insumosAdmin.forEach(i => mapInsumos[i.id] = i.nome);

    const ingredientes = dados
      .filter(r => String(r[1]) === String(idProd)) // Filtra pelo ID do Produto
      .map(r => ({
        idFicha: r[0],
        idProduto: r[1],
        // Fallback: Se o insumo foi deletado do estoque, mostra o ID original
        nomeInsumo: mapInsumos[r[2]] || r[5] || `Insumo não encontrado (${r[2]})`,
        quantidade: Number(r[3])
      }));

    return ingredientes;

  } catch (e) {
    Logger.log(`[FICHA][ERROR] Falha ao ler receita: ${e.message}`);
    throw new Error("Não foi possível carregar a receita.");
  }
}

/**
 * Remove um ingrediente da receita.
 */
function deleteIngredienteFichaTecnica(idFicha, idProd) {
  Logger.log(`[FICHA][DELETE] Removendo vínculo ID: ${idFicha}`);

  try {
    const sheet = getSheet(getPlanilha(), ABA_FICHA_TECNICA);
    const data = sheet.getDataRange().getValues();
    let found = false;

    // Loop reverso ou normal (aqui normal pois interrompemos no break)
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(idFicha)) {
        sheet.deleteRow(i + 1); // Base 1
        found = true;
        break;
      }
    }

    if (!found) throw new Error("Ingrediente não encontrado na base.");

    // Regra de Negócio: Ao mudar a receita, o custo muda. Recalcular.
    recalcularCustoDosProdutos(idProd);

    return 'Ingrediente removido e custo atualizado.';

  } catch (e) {
    Logger.log(`[FICHA][ERROR] ${e.message}`);
    throw new Error(e.message);
  }
}

/**
 * ⚙️ CORE: Recalcula o CUSTO de produção baseado nos insumos e seus preços de compra.
 * Pode ser rodado para um produto específico ou para todos (passar null).
 */
function recalcularCustoDosProdutos(idProdutoAlvo) {
  Logger.log(`[CUSTOS][CALC] Iniciando recálculo. Alvo: ${idProdutoAlvo || 'TODOS'}`);

  try {
    const ss = getPlanilha();
    const abaCardapio = getSheet(ss, ABA_CARDAPIO);
    const abaFicha = getSheet(ss, ABA_FICHA_TECNICA);
    const abaCustos = getSheet(ss, ABA_CUSTO_INSUMOS);

    // 1. Mapa de Custos dos Insumos (Custo Unitário)
    // Lê a aba Custo_Insumos para saber quanto custa cada grama/unidade
    const mapCustosInsumos = {};
    const dadosCustos = abaCustos.getDataRange().getValues();
    // Assumindo estrutura: [0]Nome, [1]ID, ... [5]CustoUnitario
    for (let i = 1; i < dadosCustos.length; i++) {
      const idInsumo = dadosCustos[i][1];
      const custoUnit = parseFloat(dadosCustos[i][5]); // Coluna F
      if (idInsumo && !isNaN(custoUnit)) {
        mapCustosInsumos[idInsumo] = custoUnit;
      }
    }

    // 2. Mapa de Receitas
    // Agrupa todos os ingredientes por produto para evitar múltiplas leituras
    const receitas = abaFicha.getDataRange().getValues();
    const mapReceitas = {}; // { ID_PROD: [ {idInsumo, qtd}, ... ] }

    for (let j = 1; j < receitas.length; j++) {
      const idProd = receitas[j][1];
      const idInsumo = receitas[j][2];
      const qtd = parseFloat(receitas[j][3]) || 0;

      if (!mapReceitas[idProd]) mapReceitas[idProd] = [];
      mapReceitas[idProd].push({ idInsumo, qtd });
    }

    // 3. Atualização do Cardápio
    const dadosCardapio = abaCardapio.getDataRange().getValues();

    for (let k = 1; k < dadosCardapio.length; k++) {
      const idProdAtual = dadosCardapio[k][0];

      // Se foi pedido recálculo de apenas um, pula os outros
      if (idProdutoAlvo && String(idProdutoAlvo) !== String(idProdAtual)) continue;

      let custoTotalProduto = 0;
      const ingredientes = mapReceitas[idProdAtual];

      if (ingredientes) {
        ingredientes.forEach(ing => {
          const custoUnitarioInsumo = mapCustosInsumos[ing.idInsumo] || 0;
          custoTotalProduto += (ing.qtd * custoUnitarioInsumo);
        });
      }

      // Atualiza a coluna CUSTO (Coluna E -> índice 4)
      // Apenas se o valor mudou para evitar escritas desnecessárias (Otimização)
      const custoAtualPlanilha = parseFloat(dadosCardapio[k][4]) || 0;

      // Usamos toFixed(2) para comparação monetária
      if (custoTotalProduto.toFixed(2) !== custoAtualPlanilha.toFixed(2)) {
        abaCardapio.getRange(k + 1, 5).setValue(custoTotalProduto);
        Logger.log(`[CUSTOS] Atualizado ${idProdAtual}: R$ ${custoTotalProduto.toFixed(2)}`);
      }
    }

  } catch (e) {
    Logger.log(`[CUSTOS][ERROR] Falha no recálculo: ${e.message}`);
    // Não lançamos throw aqui para não travar processos em lote, apenas logamos o erro
  }
}

/**
 * 📉 CORE: Baixa de Estoque Automática ao Vender.
 * Chamada quando um pedido entra.
 */
function darBaixaEstoqueFichaTecnica(ss, idProduto, qtdVenda) {
  // Logger.log(`[ESTOQUE][BAIXA] Produto: ${idProduto}, Qtd: ${qtdVenda}`);

  try {
    const abaFicha = getSheet(ss, ABA_FICHA_TECNICA);
    const abaEstoque = getSheet(ss, ABA_ESTOQUE);

    // Otimização: Ler dados apenas uma vez se possível, mas aqui lemos para garantir frescor
    const dadosFicha = abaFicha.getDataRange().getValues();

    // Filtra ingredientes do produto vendido
    const ingredientes = dadosFicha.filter(r => String(r[1]) === String(idProduto));

    if (ingredientes.length === 0) {
      // Logger.log(`[ESTOQUE][INFO] Produto ${idProduto} não tem ficha técnica. Nenhuma baixa realizada.`);
      return;
    }

    ingredientes.forEach(ing => {
      const idInsumo = ing[2];
      const qtdPorUnidade = parseFloat(ing[3]);
      const baixaTotal = qtdPorUnidade * qtdVenda;

      // Chama a função de atualização atômica do estoque
      // Passamos negativo (-) para subtrair
      atualizarEstoque(abaEstoque, idInsumo, -baixaTotal);

      // Log opcional (pode ser removido se gerar muitos dados)
      registrarLogEstoque(idInsumo, 'Venda Auto', -baixaTotal, `Venda Prod ${idProduto}`);
    });

  } catch (e) {
    Logger.log(`[ESTOQUE][ERROR] Falha na baixa de estoque: ${e.message}`);
    // Importante: Não paramos a venda se o estoque falhar, mas logamos o erro crítico.
  }
}
// ====================================================================
// 9. LÓGICA DE NEGÓCIO (Estoque, Logs, Custos)
// ====================================================================

function registrarLogEstoque(id, nome, qtd, motivo) {
  const aba = getSheet(getPlanilha(), ABA_LOG_ESTOQUE);
  aba.appendRow([new Date(), id, qtd, motivo, nome]);
}

function getEstoqueAtual(aba, id) {
  const data = aba.getDataRange().getValues();
  const row = data.find(r => r[0] === id);
  return row ? Number(row[3]) : 0;
}

function atualizarEstoque(aba, id, qtd) {
  const data = aba.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      const atual = Number(data[i][3]);
      aba.getRange(i + 1, 4).setValue(atual + qtd);
      break;
    }
  }
}

function setEstoque(aba, id, valor) {
  const data = aba.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      aba.getRange(i + 1, 4).setValue(valor);
      break;
    }
  }
}

function getNomeItemPorID(id) {
  const ss = getPlanilha();
  const data = getSheet(ss, ABA_CARDAPIO).getDataRange().getValues();
  const row = data.find(r => r[0] === id);
  return row ? row[1] : null;
}

function getNomeInsumoPorID(abaEstoque, id) {
  const data = abaEstoque.getDataRange().getValues();
  const row = data.find(r => r[0] === id);
  return row ? row[1] : null;
}





// --- Backend: Ler Número Atual ---
function getWhatsAppConfigAtual() {
  const sheet = getSheet(getPlanilha(), ABA_CONFIG);
  const data = sheet.getDataRange().getValues();
  // Procura na Coluna A pela chave 'NumeroWhatsApp'
  const row = data.find(r => r[0] === 'NumeroWhatsApp');
  // Retorna o valor da Coluna B ou vazio
  return row ? String(row[1]) : '';
}

// --- Backend: Salvar Novo Número ---
function salvarWhatsAppConfig(formData) {
  const sheet = getSheet(getPlanilha(), ABA_CONFIG);
  const data = sheet.getDataRange().getValues();

  // Sanitização: Remove tudo que não for número (espaços, traços, parenteses)
  const cleanNumber = String(formData.numero).replace(/\D/g, '');

  if (cleanNumber.length < 10) {
    throw new Error('Número inválido. Use o formato: 55 + DDD + Número (Ex: 5511999998888)');
  }

  let encontrado = false;

  // Procura a linha para atualizar
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === 'NumeroWhatsApp') {
      sheet.getRange(i + 1, 2).setValue(cleanNumber);
      encontrado = true;
      break;
    }
  }

  // Se não existir (caso raro), cria a linha
  if (!encontrado) {
    sheet.appendRow(['NumeroWhatsApp', cleanNumber]);
  }

  return `Número atualizado para: ${cleanNumber}`;
}

// ====================================================================
//  MÓDULO DE TAXAS DE ENTREGA (CRUD & API)
// ====================================================================

/**
 * UI: Abre a Sidebar de Gerenciamento de Taxas
 */
function showSidebarGerenciarTaxas() {
  const html = getHtmlTemplate('GerenciarTaxas');
  SpreadsheetApp.getUi().showSidebar(html.setTitle('Gerenciar Taxas de Entrega'));
}

/**
 * API (Public): Retorna lista de bairros para o App de Delivery.
 * Inclui opção de "A Combinar" automaticamente.
 */
function getBairrosParaDelivery() {
  const ss = getPlanilha();
  const sheet = getSheetOrCreate(ss, ABA_TAXA_ENTREGA, ['BAIRRO', 'TAXA']);
  const dados = sheet.getDataRange().getValues();
  dados.shift(); // Remove cabeçalho

  const lista = dados
    .filter(r => r[0] && r[0] !== '') // Filtra vazios
    .map(r => ({
      bairro: r[0],
      taxa: parseFloat(String(r[1]).replace(',', '.')) || 0
    }));

  // Ordena alfabeticamente
  lista.sort((a, b) => a.bairro.localeCompare(b.bairro));

  return lista;
}

/**
 * CRUD (Read): Lista taxas para o Admin
 */
function getTaxasParaAdmin() {
  return getBairrosParaDelivery(); // Reutiliza a lógica
}

/**
 * CRUD (Create/Update): Salva ou Atualiza uma taxa baseada no nome do Bairro
 */
function salvarTaxaEntrega(formData) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const ss = getPlanilha();
    const sheet = getSheetOrCreate(ss, ABA_TAXA_ENTREGA, ['BAIRRO', 'TAXA']);
    const bairroInput = String(formData.bairro).trim();
    const taxaInput = parseFloat(String(formData.taxa).replace(',', '.'));

    if (!bairroInput || isNaN(taxaInput)) throw new Error("Bairro e Taxa são obrigatórios.");

    const dados = sheet.getDataRange().getValues();
    let linhaEncontrada = -1;

    // Procura se já existe para atualizar
    for (let i = 1; i < dados.length; i++) {
      if (String(dados[i][0]).toLowerCase() === bairroInput.toLowerCase()) {
        linhaEncontrada = i + 1;
        break;
      }
    }

    if (linhaEncontrada > 0) {
      sheet.getRange(linhaEncontrada, 2).setValue(taxaInput);
      return `Taxa do bairro '${bairroInput}' atualizada para R$ ${taxaInput.toFixed(2)}`;
    } else {
      sheet.appendRow([bairroInput, taxaInput]);
      return `Bairro '${bairroInput}' cadastrado com taxa R$ ${taxaInput.toFixed(2)}`;
    }
  } finally {
    lock.releaseLock();
  }
}

/**
 * CRUD (Delete): Remove um bairro
 */
function excluirTaxaEntrega(bairro) {
  const ss = getPlanilha();
  const sheet = getSheet(ss, ABA_TAXA_ENTREGA);
  const dados = sheet.getDataRange().getValues();

  for (let i = 1; i < dados.length; i++) {
    if (String(dados[i][0]) === bairro) {
      sheet.deleteRow(i + 1);
      return `Bairro '${bairro}' removido.`;
    }
  }
  throw new Error("Bairro não encontrado.");
}

// ====================================================================
// 10. WEB APP / KDS (Delivery)
// ====================================================================

function doGet(e) {
  Logger.log("[WEBAPP] Acesso iniciado.");
  try {
    const page = e.parameter.page;
    let template;
    let title;

    if (page === 'cozinha') {
      template = 'DeliveryApp_Cozinha';
      title = '👨‍🍳 Monitor KDS';
    } else {
      template = 'DeliveryApp';
      title = 'Delivery App';
    }

    return HtmlService.createTemplateFromFile(template)
      .evaluate()
      .setTitle(title)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');

  } catch (e) {
    Logger.log(`[WEBAPP][CRITICAL] Erro ao carregar: ${e.message}`);
    return ContentService.createTextOutput("⚠️ O sistema está indisponível no momento. Por favor, avise o restaurante.")
      .setMimeType(ContentService.MimeType.TEXT);
  }
}

/**
 * Busca a taxa de entrega baseada no bairro (Case insensitive).
 * Retorna 0 se não encontrar.
 */
function getTaxaEntregaPorBairro(bairroInput) {
  if (!bairroInput) return 0;

  // Usa a função getPlanilha() que já temos
  const ss = getPlanilha();
  const aba = getSheet(ss, ABA_TAXA_ENTREGA);

  // Pega dados (evita ler se estiver vazio)
  if (aba.getLastRow() <= 1) return 0;

  const dados = aba.getRange(2, 1, aba.getLastRow() - 1, 2).getValues();
  const bairroLimpo = String(bairroInput).trim().toLowerCase();

  for (let i = 0; i < dados.length; i++) {
    const bairroBanco = String(dados[i][0]).trim().toLowerCase(); // Coluna A
    const taxa = parseFloat(dados[i][1]); // Coluna B

    // Verifica igualdade ou se contém (Ex: "Centro" bate com "Centro Histórico")
    if (bairroBanco && (bairroBanco === bairroLimpo || bairroLimpo.includes(bairroBanco))) {
      return isNaN(taxa) ? 0 : taxa;
    }
  }
  return 0; // Se não achar, taxa zero (A combinar)
}

function getCardapioParaDelivery() {
  const ss = getPlanilha();
  const cardapio = getCardapioParaAdmin();
  const linksSheet = getSheet(ss, ABA_PRODUTO_ADICIONAIS_LINK);
  const linksData = linksSheet.getDataRange().getValues();
  const adicionais = getAdicionaisParaAdmin();

  // Mapeia adicionais
  const addMap = {};
  adicionais.forEach(a => { if (a.ativo === 'SIM') addMap[a.id] = a; });

  const grouped = {};

  cardapio.forEach(item => {
    if (item.ativo !== 'SIM') return;

    if (!grouped[item.categoria]) grouped[item.categoria] = [];

    // Busca adicionais deste produto
    const meusLinks = linksData.filter(l => String(l[0]) === String(item.id) && l[2] === 'SIM');
    const meusAdds = meusLinks.map(l => addMap[l[1]]).filter(a => a); // remove undefined

    grouped[item.categoria].push({
      id: item.id,
      nome: item.nome,
      preco: item.preco,
      descricao: item.descricao,
      foto_url: item.foto_url,
      adicionaisPermitidos: meusAdds
    });
  });

  return grouped;
}


function processarPedidoDelivery(jsonPedido) {
  const pedido = JSON.parse(jsonPedido);
  const ss = getPlanilha();
  const abaVendas = getSheet(ss, ABA_VENDAS);
  const idVenda = `VD-${Date.now()}`;

  // 1. Calcula Total dos Produtos
  let totalProdutos = 0;
  pedido.itens.forEach(item => {
    // Baixa Estoque
    darBaixaEstoqueFichaTecnica(ss, item.id, item.qtde);

    let itemTotal = item.precoBase;
    if (item.adicionais) {
      item.adicionais.forEach(a => itemTotal += a.preco);
    }
    totalProdutos += (itemTotal * item.qtde);
  });

  // 2. Calcula Taxa de Entrega (Backend Authority)
  const taxaEntrega = getTaxaEntregaPorBairro(pedido.bairro);

  // 3. Total Final
  const totalFinal = totalProdutos + taxaEntrega;

  // 4. Grava na Planilha
  abaVendas.appendRow([
    idVenda,
    new Date(),
    jsonPedido,
    totalFinal, // Grava o Total COM a taxa
    pedido.pagamento,
    pedido.clienteNome,
    pedido.clienteTelefone,
    pedido.logradouro,
    pedido.numero,
    '',
    pedido.bairro,
    '',
    pedido.referencia,
    taxaEntrega, // Grava a taxa na coluna N (se sua estrutura permitir) ou ignora
    'NOVO',
    ''
  ]);

  // 5. Gera Mensagem passando os valores separados
  const msg = gerarMensagemWhatsApp(ss, idVenda, pedido, totalProdutos, taxaEntrega, new Date());

  return { status: 'sucesso', idVenda, mensagemWhatsApp: msg };
}





/**
 * Gera o link do WhatsApp com detalhamento de preços por item.
 */
/**
 * Gera o link do WhatsApp usando Unicode Escapes para garantir que os Emojis
 * apareçam em qualquer dispositivo (Android, iOS, Web) sem erros de codificação.
 */
function gerarMensagemWhatsApp(ss, idVenda, pedido, totalProdutos, taxaEntrega, dataHora) {
  Logger.log(`>>> gerarMensagemWhatsApp: Gerando mensagem Unicode para venda ${idVenda}.`);

  // --- MAPA DE EMOJIS SEGUROS (UNICODE) ---
  const EMOJI = {
    ROBO: '\uD83E\uDD16',      // 🤖
    BURGER: '\uD83C\uDF54',    // 🍔
    ID: '\uD83C\uDD94',        // 🆔
    USER: '\uD83D\uDC64',      // 👤
    MEMO: '\uD83D\uDCDD',      // 📝
    PIN: '\uD83D\uDCCD',       // 📍
    ROCKET: '\uD83D\uDE80',    // 🚀
    HOUSE: '\uD83C\uDFE0',     // 🏠
    PHONE: '\uD83D\uDCDE',     // 📞
    MONEY: '\uD83D\uDCB0',     // 💰
    SCOOTER: '\uD83D\uDEF5',   // 🛵
    FLY_MONEY: '\uD83D\uDCB8', // 💸
    CARD: '\uD83D\uDCB3',      // 💳
    CLOCK: '\uD83D\uDD52',     // 🕒
    CHECK: '\u2705',           // ✅
    WARNING: '\u26A0\uFE0F',   // ⚠️
    BULLET: '\u25AA\uFE0F',    // ▪️
    ARROW: '\u2937'            // ⤵
  };

  // Fallback de segurança para SS
  if (!ss || typeof ss.getSheetByName !== 'function') {
    try { ss = SpreadsheetApp.openById(ID_PLANILHA_MESTRA); } catch (e) { }
  }

  try {
    const configSheet = ss.getSheetByName(ABA_CONFIG);
    const config = configSheet ? configSheet.getDataRange().getValues() : [];
    const numeroConfig = config.find(row => row[0] === 'NumeroWhatsApp');
    let numero = numeroConfig ? String(numeroConfig[1]).replace(/\D/g, '') : '5511999998888';

    const dataFormatada = dataHora.toLocaleDateString('pt-BR');
    const horaFormatada = dataHora.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });

    // --- MONTAGEM DA LISTA DE ITENS ---
    let listaItens = '';

    (pedido.itens || []).forEach(item => {
      const precoUnitario = parseFloat(item.precoBase || item.preco || 0);

      // Item Principal
      listaItens += `${EMOJI.BULLET} *${item.qtde}x ${item.nome}* ... R$ ${precoUnitario.toFixed(2)}\n`;

      // Adicionais
      if (Array.isArray(item.adicionais) && item.adicionais.length > 0) {
        item.adicionais.forEach(adicional => {
          const precoAdd = parseFloat(adicional.preco || 0);
          listaItens += `   ${EMOJI.ARROW} _+ ${adicional.nome}_ ... R$ ${precoAdd.toFixed(2)}\n`;
        });
      }

      // Obs
      if (item.obs) listaItens += `   ${EMOJI.WARNING} _Obs: ${item.obs}_\n`;

      listaItens += `\n`;
    });

    // Cálculos
    const totalGeral = totalProdutos + taxaEntrega;

    // --- CORPO DA MENSAGEM ---
    let msg = `${EMOJI.ROBO} *NOVO PEDIDO REALIZADO!* ${EMOJI.BURGER}\n`;
    msg += `${EMOJI.ID} *ID:* ${idVenda}\n`;
    msg += `${EMOJI.USER} *Cliente:* ${pedido.clienteNome}\n`;
    msg += `───────────────────\n\n`;

    msg += `${EMOJI.MEMO} *RESUMO DO PEDIDO*\n\n`;
    msg += listaItens;
    msg += `───────────────────\n\n`;

    msg += `${EMOJI.PIN} *ENTREGA PARA:*\n`;
    msg += `${EMOJI.ROCKET} *RUA / AV:*${pedido.logradouro}, ${pedido.numero}\n`;
    msg += `${EMOJI.HOUSE} Bairro: *${pedido.bairro}*\n`;
    if (pedido.referencia) msg += `👀 Ref: ${pedido.referencia}\n`;
    msg += `${EMOJI.PHONE} *Tel:* ${pedido.clienteTelefone}\n\n`;

    msg += `${EMOJI.MONEY} *VALORES:*\n`;
    msg += `Subtotal: R$ ${parseFloat(totalProdutos).toFixed(2)}\n`;

    if (taxaEntrega > 0) {
      msg += `${EMOJI.SCOOTER} *Taxa Entrega:* R$ ${parseFloat(taxaEntrega).toFixed(2)}\n`;
    } else {
      msg += `${EMOJI.SCOOTER} *Taxa Entrega:* *A Combinar*\n`;
    }

    msg += `${EMOJI.FLY_MONEY} *TOTAL GERAL: R$ ${totalGeral.toFixed(2)}*\n`;
    msg += `${EMOJI.CARD} Pagamento: *${pedido.pagamento}*\n`;
    msg += `───────────────────\n`;

    msg += `${EMOJI.CLOCK} *Gerado em:* ${dataFormatada} às ${horaFormatada}\n`;
    msg += `${EMOJI.CHECK} _Aguardando confirmação do restaurante._`;

    return `https://wa.me/${numero}?text=${encodeURIComponent(msg)}`;

  } catch (err) {
    Logger.log(`*** ERRO msg: ${err.message}`);
    return `https://wa.me/5511999998888?text=ErroAoGerarLink`;
  }
}

// ====================================================================
// ===> MÓDULO KDS (MONITOR DE COZINHA)
// ====================================================================

/**
 * Busca pedidos das últimas 24h para o Monitor.
 * Versão Corrigida: Compatível com JSON Completo ou Lista Simples.
 */
function getPedidosCozinha() {
  const ss = getPlanilha();
  const aba = getSheet(ss, ABA_VENDAS);

  if (aba.getLastRow() <= 1) return [];

  const dados = aba.getDataRange().getDisplayValues();
  dados.shift(); // Remove cabeçalho

  const agora = new Date();
  const limiteHoras = 24 * 60 * 60 * 1000; // 24h

  const pedidos = dados
    .map((row, index) => {
      // 1. Parser de Data (Mantido)
      const dataString = row[1];
      let dataPedido;
      if (dataString.includes('/')) {
        const partes = dataString.split(' ');
        const dataPartes = partes[0].split('/');
        const horaPartes = partes[1] ? partes[1].split(':') : [0, 0, 0];
        dataPedido = new Date(dataPartes[2], dataPartes[1] - 1, dataPartes[0], horaPartes[0], horaPartes[1], horaPartes[2] || 0);
      } else {
        dataPedido = new Date(dataString);
      }

      // 2. Filtro de Tempo/Status (Mantido)
      const isRecent = (agora - dataPedido < limiteHoras);
      const isPending = (row[14] !== 'ENTREGUE' && row[14] !== 'CANCELADO');
      if (!isRecent && !isPending) return null;

      // 3. CORREÇÃO CRÍTICA: Extração dos Itens
      let rawJson = null;
      let listaItens = [];

      try {
        rawJson = JSON.parse(row[2]);
      } catch (e) {
        rawJson = [];
      }

      if (rawJson) {
        if (Array.isArray(rawJson)) {
          // Formato Antigo: O JSON é direto a lista de itens
          listaItens = rawJson;
        } else if (rawJson.itens && Array.isArray(rawJson.itens)) {
          // Formato Novo: O JSON é o pedido completo, pegamos a propriedade .itens
          listaItens = rawJson.itens;
        }
      }

      return {
        rowIndex: index + 2,
        id: row[0],
        dataHora: dataPedido.toISOString(),
        itens: listaItens, // Agora garantimos que é sempre uma lista!
        total: row[3],
        pagamento: row[4],
        clienteNome: row[5],
        clienteTel: row[6],
        logradouro: row[7],
        numero: row[8],
        bairro: row[10],
        status: row[14],
        obs: row[15]
      };
    })
    .filter(p => p !== null)
    .sort((a, b) => new Date(a.dataHora) - new Date(b.dataHora));

  return pedidos;
}

/**
 * Atualiza o Status e executa lógicas de borda (Notificação ou Estorno).
 */
function atualizarStatusPedido(idPedido, novoStatus) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    throw new Error('O sistema está ocupado. Tente novamente.');
  }

  try {
    const ss = getPlanilha();
    const aba = getSheet(ss, ABA_VENDAS);
    const dados = aba.getDataRange().getValues();

    let linha = -1;
    let clienteTel = '';
    let clienteNome = '';
    let jsonPedidoOriginal = '';
    let statusAtual = '';

    // Busca o pedido
    for (let i = 1; i < dados.length; i++) {
      if (String(dados[i][0]) === String(idPedido)) {
        linha = i + 1;
        jsonPedidoOriginal = dados[i][2]; // Coluna C (JSON)
        clienteNome = dados[i][5];
        clienteTel = dados[i][6];
        statusAtual = dados[i][14]; // Coluna O
        break;
      }
    }

    if (linha === -1) throw new Error(`Pedido ${idPedido} não encontrado.`);

    // Proteção: Não cancelar duas vezes
    if (statusAtual === 'CANCELADO' && novoStatus === 'CANCELADO') {
      return { success: true, message: 'Pedido já estava cancelado.' };
    }

    // 1. Grava o Novo Status
    aba.getRange(linha, 15).setValue(novoStatus);

    // 2. Lógica de Estorno (Se for cancelamento)
    if (novoStatus === 'CANCELADO') {
      estornarEstoquePedido(ss, jsonPedidoOriginal, idPedido);
    }

    SpreadsheetApp.flush();

    // 3. Lógica de Notificação WhatsApp
    let linkZap = null;
    if (clienteTel && (novoStatus === 'PREPARANDO' || novoStatus === 'SAIU_ENTREGA')) {
      const num = String(clienteTel).replace(/\D/g, '');
      let texto = '';

      if (novoStatus === 'PREPARANDO') {
        texto = `Olá ${clienteNome}! 👨‍🍳 Seu pedido *${idPedido}* começou a ser preparado.`;
      } else if (novoStatus === 'SAIU_ENTREGA') {
        texto = `🛵 Saiu para entrega! O pedido *${idPedido}* está a caminho.`;
      }

      if (num.length >= 10) {
        linkZap = `https://wa.me/55${num}?text=${encodeURIComponent(texto)}`;
      }
    }

    return { success: true, linkNotificacao: linkZap };

  } catch (e) {
    Logger.log(`Erro ao atualizar status: ${e.message}`);
    throw new Error(e.message);
  } finally {
    lock.releaseLock();
  }
}

// ====================================================================
// 10. MÓDULO PAGAMENTOS (CRUD DINÂMICO)
// ====================================================================

/**
 * UI: Abre Sidebar
 */
function showSidebarGerenciarPagamentos() {
  const html = getHtmlTemplate('GerenciarPagamentos');
  SpreadsheetApp.getUi().showSidebar(html.setTitle('Formas de Pagamento'));
}

/**
 * Adicione ao MENU 'onOpen':
 * .addItem('💳 Formas de Pagamento', 'showSidebarGerenciarPagamentos')
 */

/**
 * API: Retorna pagamentos ativos para o App
 */
function getPagamentosParaDelivery() {
  const ss = getPlanilha();
  const sheet = getSheetOrCreate(ss, ABA_PAGAMENTOS, ['METODO', 'ATIVO']);
  const dados = sheet.getDataRange().getValues();
  dados.shift(); // Remove header

  // Se estiver vazio, retorna padrão para não quebrar o app
  if (dados.length === 0) return ['PIX', 'Dinheiro', 'Cartão de Crédito', 'Cartão de Débito'];

  return dados
    .filter(r => r[0] && String(r[1]) === 'SIM') // Apenas Ativos
    .map(r => r[0]);
}

/**
 * CRUD: Lista para Admin
 */
function getPagamentosAdmin() {
  const ss = getPlanilha();
  const sheet = getSheetOrCreate(ss, ABA_PAGAMENTOS, ['METODO', 'ATIVO']);
  const dados = sheet.getDataRange().getValues();
  dados.shift();
  return dados.map(r => ({ metodo: r[0], ativo: r[1] })).filter(r => r.metodo);
}

/**
 * CRUD: Salvar Novo
 */
function salvarPagamento(formData) {
  const ss = getPlanilha();
  const sheet = getSheet(ss, ABA_PAGAMENTOS);
  const metodo = String(formData.metodo).trim();

  if (!metodo) throw new Error("Digite o nome do método.");

  // Verifica duplicidade
  const dados = sheet.getDataRange().getValues();
  for (let i = 1; i < dados.length; i++) {
    if (String(dados[i][0]).toLowerCase() === metodo.toLowerCase()) {
      throw new Error("Este método já existe.");
    }
  }

  sheet.appendRow([metodo, 'SIM']);
  return 'Método de pagamento adicionado!';
}

/**
 * CRUD: Excluir (Físico)
 */
function excluirPagamento(metodo) {
  const ss = getPlanilha();
  const sheet = getSheet(ss, ABA_PAGAMENTOS);
  const dados = sheet.getDataRange().getValues();

  for (let i = 1; i < dados.length; i++) {
    if (String(dados[i][0]) === metodo) {
      sheet.deleteRow(i + 1);
      return 'Método removido.';
    }
  }
  throw new Error("Método não encontrado.");
}

// ====================================================================
// 9. MÓDULO BI & FINANCEIRO (DASHBOARD AUTOMÁTICO)
// ====================================================================

/**
 * UI: Botão para atualizar o Dashboard manualmente via Menu
 */
function uiAtualizarDashboard() {
  const ui = SpreadsheetApp.getUi();
  try {
    gerarDashboardFinanceiro();
    ui.alert('📊 Dashboard atualizado com sucesso!');
  } catch (e) {
    ui.alert('Erro ao gerar dashboard: ' + e.message);
  }
}

/**
 * Adicione esta linha no seu MENU 'onOpen' existente (dentro de Estoque & Custos):
 * .addItem('📊 Atualizar Dashboard', 'uiAtualizarDashboard')
 */

/**
 * CORE: Gera a aba DASHBOARD com métricas calculadas e design limpo.
 */
function gerarDashboardFinanceiro() {
  const ss = getPlanilha();
  let sheet = ss.getSheetByName('DASHBOARD');

  // Se não existir, cria e move para o início (posição 1)
  if (!sheet) {
    sheet = ss.insertSheet('DASHBOARD', 0);
  } else {
    sheet.clear(); // Limpa dados antigos para regenerar
  }

  // --- 1. Coleta e Processamento de Dados ---
  const abaVendas = getSheet(ss, ABA_VENDAS);
  if (abaVendas.getLastRow() <= 1) {
    sheet.getRange("A1").setValue("Sem vendas registradas para gerar indicadores.");
    return;
  }

  const dados = abaVendas.getDataRange().getValues();
  dados.shift(); // Remove header

  // Variáveis de Acumulação
  const hoje = new Date();
  hoje.setHours(0, 0, 0, 0);

  let fatHoje = 0;
  let fatTotal = 0;
  let qtdPedidosHoje = 0;
  let qtdPedidosTotal = dados.length;
  let mapPagamentos = {};
  let mapProdutos = {};

  const statusPedido = row[14]; // Coluna O
  if (statusPedido === 'CANCELADO') return;

  dados.forEach(row => {
    // Mapeamento das colunas (Ajuste se mudou a ordem)
    // Coluna B (1) = Data, D (3) = Total, E (4) = Pagamento, C (2) = JSON
    const dataVenda = new Date(row[1]);
    const valor = parseFloat(row[3]) || 0; // Coluna D
    const pag = row[4] || 'Não Informado'; // Coluna E

    // Totais Gerais
    fatTotal += valor;

    // Acumula por Tipo de Pagamento
    mapPagamentos[pag] = (mapPagamentos[pag] || 0) + valor;

    // Filtro HOJE
    const dataCheck = new Date(dataVenda);
    dataCheck.setHours(0, 0, 0, 0);
    if (dataCheck.getTime() === hoje.getTime()) {
      fatHoje += valor;
      qtdPedidosHoje++;
    }

    // Contagem de Produtos (Extração Inteligente do JSON)
    try {
      let itens = [];
      try { itens = JSON.parse(row[2]); } catch (e) { }

      // Normaliza formato (Lista antiga ou Objeto novo)
      if (!Array.isArray(itens) && itens.itens) itens = itens.itens;

      if (Array.isArray(itens)) {
        itens.forEach(i => {
          const nome = i.nome || 'Item Desconhecido';
          const qtd = parseInt(i.qtde) || 1;
          mapProdutos[nome] = (mapProdutos[nome] || 0) + qtd;
        });
      }
    } catch (e) { }
  });

  // --- 2. Renderização Visual (Design) ---

  // Estilos
  const styleTitulo = SpreadsheetApp.newTextStyle().setFontSize(16).setBold(true).setForegroundColor("#111827").build();
  const styleLabel = SpreadsheetApp.newTextStyle().setFontSize(10).setForegroundColor("#6b7280").setBold(true).build();
  const styleValor = SpreadsheetApp.newTextStyle().setFontSize(22).setBold(true).setForegroundColor("#059669").build(); // Verde
  const styleValorSec = SpreadsheetApp.newTextStyle().setFontSize(22).setBold(true).setForegroundColor("#3b82f6").build(); // Azul

  // Cabeçalho
  sheet.getRange("B2").setValue("🚀 Painel de Controle Financeiro").setTextStyle(styleTitulo);
  sheet.getRange("B3").setValue("Última atualização: " + new Date().toLocaleString());

  // --- CARDS KPI (Linha 5) ---

  // Card 1: Vendas Hoje
  sheet.getRange("B5").setValue("FATURAMENTO HOJE").setTextStyle(styleLabel);
  sheet.getRange("B6").setValue(fatHoje).setNumberFormat("R$ #,##0.00").setTextStyle(styleValor);

  // Card 2: Pedidos Hoje
  sheet.getRange("D5").setValue("PEDIDOS HOJE").setTextStyle(styleLabel);
  sheet.getRange("D6").setValue(qtdPedidosHoje).setNumberFormat("0").setTextStyle(styleValorSec);

  // Card 3: Ticket Médio (Geral)
  const ticketMedio = qtdPedidosTotal > 0 ? (fatTotal / qtdPedidosTotal) : 0;
  sheet.getRange("F5").setValue("TICKET MÉDIO (GERAL)").setTextStyle(styleLabel);
  sheet.getRange("F6").setValue(ticketMedio).setNumberFormat("R$ #,##0.00").setTextStyle(styleValor);

  // --- TABELAS (Linha 9) ---

  // Tabela 1: Top Produtos
  sheet.getRange("B9").setValue("🏆 Top 5 Produtos Mais Vendidos").setFontWeight("bold").setFontSize(11);
  sheet.getRange("B10:C10").setValues([["Produto", "Qtd Vendida"]]).setBackground("#f3f4f6").setFontWeight("bold");

  // Ordena e Pega Top 5
  const sortedProdutos = Object.entries(mapProdutos)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5);

  if (sortedProdutos.length > 0) {
    sheet.getRange(11, 2, sortedProdutos.length, 2).setValues(sortedProdutos);
  } else {
    sheet.getRange("B11").setValue("Sem dados de produtos.");
  }

  // Tabela 2: Faturamento por Pagamento
  sheet.getRange("E9").setValue("💳 Receita por Pagamento").setFontWeight("bold").setFontSize(11);
  sheet.getRange("E10:F10").setValues([["Método", "Total (R$)"]]).setBackground("#f3f4f6").setFontWeight("bold");

  const sortedPag = Object.entries(mapPagamentos).sort((a, b) => b[1] - a[1]);
  if (sortedPag.length > 0) {
    sheet.getRange(11, 5, sortedPag.length, 2).setValues(sortedPag);
    sheet.getRange(11, 6, sortedPag.length, 1).setNumberFormat("R$ #,##0.00");
  }

  // --- Ajustes Finais de Layout ---
  sheet.setColumnWidth(1, 20); // Margem A
  sheet.setColumnWidth(2, 220); // Coluna B (Card 1 / Prod Nome)
  sheet.setColumnWidth(3, 50);  // Espaço
  sheet.setColumnWidth(4, 180); // Coluna D (Card 2)
  sheet.setColumnWidth(5, 150); // Coluna E (Pagamento Nome)
  sheet.setColumnWidth(6, 120); // Coluna F (Card 3 / Pag Valor)

  sheet.setHiddenGridlines(true);
}


/**
 * SETUP: Executar esta função UMA VEZ para ligar a atualização automática.
 * Configura o Dashboard para atualizar sozinho a cada 1 hora.
 */
function setupAtualizacaoAutomatica() {
  const nomeFuncao = 'gerarDashboardFinanceiro';

  // 1. Remove agendamentos antigos para não duplicar
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === nomeFuncao) {
      ScriptApp.deleteTrigger(t);
    }
  });

  // 2. Cria novo agendamento (A cada 1 hora)
  ScriptApp.newTrigger(nomeFuncao)
    .timeBased()
    .everyHours(1)
    .create();

  SpreadsheetApp.getUi().alert('✅ Robô ativado! O Dashboard será atualizado a cada 1 hora automaticamente.');
}

/**
 * Realiza o estorno (devolução) dos itens ao estoque quando um pedido é cancelado.
 */
function estornarEstoquePedido(ss, jsonPedido, idVenda) {
  Logger.log(`[ESTORNO] Iniciando estorno do pedido ${idVenda}`);

  let pedido;
  try {
    pedido = JSON.parse(jsonPedido);
  } catch (e) {
    Logger.log(`[ERRO ESTORNO] JSON inválido: ${e.message}`);
    return;
  }

  // Normaliza lista de itens (trata formato antigo/novo)
  let listaItens = [];
  if (Array.isArray(pedido)) listaItens = pedido;
  else if (pedido.itens && Array.isArray(pedido.itens)) listaItens = pedido.itens;

  const abaFicha = getSheet(ss, ABA_FICHA_TECNICA);
  const abaEstoque = getSheet(ss, ABA_ESTOQUE);
  const dadosFicha = abaFicha.getDataRange().getValues();

  // Mapa para converter ID do Insumo em Nome (para o Log)
  const dadosEstoque = abaEstoque.getDataRange().getValues();
  const mapaNomes = {};
  for (let k = 1; k < dadosEstoque.length; k++) {
    mapaNomes[dadosEstoque[k][0]] = dadosEstoque[k][1];
  }

  listaItens.forEach(item => {
    // Filtra ingredientes deste produto na ficha técnica
    const ingredientes = dadosFicha.filter(r => String(r[1]) === String(item.id));

    if (ingredientes.length > 0) {
      ingredientes.forEach(ing => {
        const idInsumo = ing[2]; // Coluna C: ID Insumo
        const qtdReceita = Number(ing[3]); // Coluna D: Qtd Usada
        const qtdEstornar = qtdReceita * item.qtde; // Devolve Qtde unitária * Qtde vendida

        // Atualiza Estoque (Positivo = Entrada/Devolução)
        atualizarEstoque(abaEstoque, idInsumo, qtdEstornar);

        // Registra Log
        const nomeInsumo = mapaNomes[idInsumo] || 'Insumo';
        registrarLogEstoque(idInsumo, nomeInsumo, qtdEstornar, `Estorno Cancelamento ${idVenda}`);
      });
    } else {
      // Se for produto 1:1 (sem ficha), tenta estornar pelo ID do produto se ele existir no estoque
      // (Opcional, depende se você cadastrou produtos de revenda no estoque com mesmo ID)
    }
  });
  Logger.log(`[ESTORNO] Concluído com sucesso.`);
}

// ====================================================================
// [NOVO] BACKEND PARA MOTOQUEIROS & BALCÃO (CRUDs EXTRAS)
// ====================================================================

// --- MOTOQUEIROS ---

function getMotoqueirosAdmin() {
  try {
    const sheet = getSheetOrCreate(getPlanilha(), SHEET_NAMES.MOTOQUEIROS, ['ID', 'NOME', 'TELEFONE', 'PLACA', 'ATIVO']);
    if (sheet.getLastRow() <= 1) return [];
    const data = sheet.getDataRange().getValues();
    data.shift();
    return data.map((r, i) => ({
       id: r[0], nome: r[1], telefone: r[2], placa: r[3], ativo: r[4]
    })).filter(x => x.id && x.ativo === 'SIM');
  } catch(e) { return []; }
}

function salvarMotoqueiro(data) {
  try {
    if(!data.nome) throw new Error("Nome obrigatório.");
    const sheet = getSheet(getPlanilha(), SHEET_NAMES.MOTOQUEIROS);
    const id = `MOT-${Date.now()}`;
    sheet.appendRow([id, data.nome, data.telefone, data.placa, 'SIM']);
    return "Motoqueiro cadastrado!";
  } catch(e) { throw new Error(e.message); }
}

function excluirMotoqueiro(id) {
  try {
    const sheet = getSheet(getPlanilha(), SHEET_NAMES.MOTOQUEIROS);
    const data = sheet.getDataRange().getValues();
    for(let i=1; i<data.length; i++) {
       if(String(data[i][0]) === String(id)) {
          sheet.deleteRow(i+1);
          return "Removido.";
       }
    }
    throw new Error("Não encontrado.");
  } catch(e) { throw new Error(e.message); }
}

// --- VENDA BALCÃO (REUTILIZA PROCESSAR PEDIDO) ---
// A função 'processarPedidoDelivery' já é genérica o suficiente.
// O Frontend do Balcão apenas chama ela passando um JSON com "bairro: Balcão".

// ====================================================================
// 11. HTML TEMPLATES (FÁBRICA - VERSÃO FINAL COM BALCÃO E MOTOQUEIROS)
// ====================================================================

function getHtmlTemplate(templateName) {
  let html = '';

  // --- 1. HEAD E ESTILOS GERAIS ---
  const head = `
    <head>
      <base target="_top">
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <script src="https://cdn.tailwindcss.com"></script>
      <link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet">
      <style>
        body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; padding: 1rem; background-color: #f9fafb; }
        .loader { border: 3px solid #f3f3f3; border-top: 3px solid #4f46e5; border-radius: 50%; width: 24px; height: 24px; animation: spin 1s linear infinite; margin: 0 auto; }
        @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
        .checkbox-label { display: flex; align-items: center; padding: 0.5rem; border: 1px solid #e5e7eb; border-radius: 0.375rem; cursor: pointer; transition: background 0.2s; margin-bottom: 0.5rem; }
        .checkbox-label:hover { background-color: #f3f4f6; }
        .form-checkbox { height: 1.25rem; width: 1.25rem; margin-right: 0.5rem; color: #4f46e5; border-radius: 4px; }
        .btn-delete { color: #ef4444; font-weight: bold; cursor: pointer; padding: 2px 6px; border-radius: 4px; }
        .btn-delete:hover { background-color: #fee2e2; }
        .modal { display: none; position: fixed; inset: 0; background-color: rgba(0,0,0,0.5); align-items: center; justify-content: center; z-index: 50; }
        .modal.show { display: flex; }
      </style>
    </head>
  `;

  // --- 2. GERADOR DE SCRIPT PADRÃO ---
  const getScriptForm = (funcName) => `
    <script>
      function s(l){const x=document.getElementById('loader');if(x)x.style.display=l?'block':'none';const b=document.getElementById('btn-submit');if(b)b.disabled=l}
      function m(t,e){const x=document.getElementById('msg');if(x){x.textContent=t;x.className=e?'text-red-600 text-center font-bold':'text-green-600 text-center font-bold'}}
      
      const form = document.getElementById('main-form');
      if(form) {
          form.addEventListener('submit', function(e) {
            e.preventDefault();
            s(true);
            
            const formData = {};
            new FormData(form).forEach((v, k) => formData[k] = v);

            // Lógica Especial para Checkboxes
            if('${funcName}' === 'processarLinkAdicionais') {
               const checks = document.querySelectorAll('.chk-addon:checked');
               formData.adicionais = Array.from(checks).map(c => c.value);
            }

            google.script.run
              .withSuccessHandler(res => {
                 s(false); m(res);
                 if('${funcName}' !== 'processarLinkAdicionais') form.reset();
                 // Hooks de Recarga
                 if(typeof loadData === 'function') loadData();
                 if(typeof loadTaxas === 'function') loadTaxas();
                 if(typeof loadPags === 'function') loadPags();
                 if(typeof loadMotos === 'function') loadMotos(); // Hook Motoqueiros
                 if(typeof carregarFichaDoItem === 'function') carregarFichaDoItem();
              })
              .withFailureHandler(err => { s(false); m(err.message, true); })
              .${funcName}(formData);
          });
      }
    </script>
  `;

  // --- 3. ROTEAMENTO DE TEMPLATES ---
  
  // >>> TELA: Cadastro Insumo
  if (templateName === 'CadastroInsumo') {
    html = `<html>${head}<body><h3 class="font-bold text-lg mb-4 text-gray-800">Novo Insumo</h3><form id="main-form" class="space-y-3"><input name="nome" placeholder="Nome" class="w-full border p-2 rounded" required><input name="unidade" placeholder="Unidade (KG, UN)" class="w-full border p-2 rounded" required><input name="estoque" type="number" placeholder="Estoque Inicial" class="w-full border p-2 rounded"><button id="btn-submit" class="w-full bg-indigo-600 text-white p-2 rounded hover:bg-indigo-700">Salvar</button></form><div id="loader" class="loader hidden mt-4"></div><div id="msg" class="mt-2"></div>${getScriptForm('processarCadastroInsumo')}</body></html>`;
  }

  // >>> TELA: Cadastro Cardápio
  else if (templateName === 'CadastroCardapio') {
    html = `<html>${head}<body><h3 class="font-bold text-lg mb-4 text-gray-800">Novo Item Cardápio</h3><form id="main-form" class="space-y-3"><input name="nome" placeholder="Nome do Item" required class="w-full border p-2 rounded"><input name="categoria" placeholder="Categoria" required class="w-full border p-2 rounded"><input name="preco" type="number" step="0.01" placeholder="Preço Venda (R$)" required class="w-full border p-2 rounded"><textarea name="descricao" placeholder="Descrição" class="w-full border p-2 rounded"></textarea><input name="foto_url" placeholder="URL da Foto" class="w-full border p-2 rounded"><select name="ativo" class="w-full border p-2 rounded"><option value="SIM">Ativo</option><option value="NAO">Inativo</option></select><button id="btn-submit" class="w-full bg-indigo-600 text-white p-2 rounded hover:bg-indigo-700">Salvar</button></form><div id="loader" class="loader hidden mt-4"></div><div id="msg" class="mt-2"></div>${getScriptForm('processarCadastroCardapio')}</body></html>`;
  }

  // >>> TELA: Gerenciar Cardápio
  else if (templateName === 'GerenciarCardapio') {
    const scriptCRUD = `<script>
        let allProducts=[]; window.addEventListener('load', loadProducts);
        function showSidebarLoading(l){const el=document.getElementById('sidebar-loader');if(el)el.style.display=l?'block':'none'}
        function loadProducts(){showSidebarLoading(true);google.script.run.withSuccessHandler(onProductsLoaded).withFailureHandler(onSidebarError).getCardapioParaAdmin()}
        function onProductsLoaded(p){allProducts=p;renderList(p);showSidebarLoading(false)}
        function onSidebarError(e){showSidebarLoading(false);document.getElementById('product-list').innerHTML='<p class="text-red-500">'+e.message+'</p>'}
        function renderList(p){const l=document.getElementById('product-list');l.innerHTML='';if(p.length===0){l.innerHTML='<p class="text-gray-500">Vazio.</p>';return}p.forEach(x=>{const c=x.ativo==='SIM'?'text-green-600':'text-red-600';l.innerHTML+=\`<div class="bg-white p-3 border rounded shadow-sm mb-2 flex justify-between items-center"><div><div class="font-bold text-gray-800">\${x.nome}</div><div class="text-xs text-gray-500">\${x.categoria} | R$ \${x.preco.toFixed(2)}</div><div class="text-xs \${c} font-bold">\${x.ativo}</div></div><button onclick="openEditModal('\${x.id}')" class="text-indigo-600 border border-indigo-200 px-2 py-1 rounded hover:bg-indigo-50 font-medium text-sm">Editar</button></div>\`})}
        document.getElementById('search-box').addEventListener('input',e=>{const t=e.target.value.toLowerCase();renderList(allProducts.filter(p=>p.nome.toLowerCase().includes(t)))});
        const modal=document.getElementById('edit-modal');
        function closeModal(){modal.classList.remove('show')}
        function openEditModal(id){const p=allProducts.find(x=>x.id===id);if(!p)return;document.getElementById('modal-id').value=p.id;document.getElementById('modal-rowIndex').value=p.rowIndex;document.getElementById('modal-nome').value=p.nome;document.getElementById('modal-preco').value=p.preco;document.getElementById('modal-categoria').value=p.categoria;document.getElementById('modal-descricao').value=p.descricao||'';document.getElementById('modal-foto_url').value=p.foto_url||'';document.getElementById('modal-ativo').value=p.ativo;document.getElementById('modal-msg').textContent='';modal.classList.add('show')}
        document.getElementById('modal-form').addEventListener('submit',e=>{e.preventDefault();const btn=document.getElementById('modal-submit-btn');btn.disabled=true;btn.textContent='Salvando...';const d={};new FormData(e.target).forEach((v,k)=>d[k]=v);google.script.run.withSuccessHandler(r=>{document.getElementById('modal-msg').textContent=r;document.getElementById('modal-msg').className='text-green-600 mt-2 text-center font-bold';btn.disabled=false;btn.textContent='Salvar Alterações';loadProducts();setTimeout(closeModal,1000)}).updateProdutoCardapio(d)});
      </script>`;
    html = `<html>${head}<body class="p-0 bg-gray-50">
      <div id="sidebar-loader" class="loader" style="display: block; margin: 2rem auto;"></div>
      <div class="p-3 bg-white border-b sticky top-0 z-10 shadow-sm"><h3 class="text-lg font-bold text-gray-800 mb-2">Gerenciar Cardápio</h3><input type="text" id="search-box" class="w-full border p-2 rounded bg-gray-50 focus:bg-white" placeholder="Buscar..."></div>
      <div id="product-list" class="p-2 space-y-2"></div>
      <div id="edit-modal" class="modal"><div class="bg-white p-4 rounded w-11/12 max-w-md max-h-[90vh] overflow-y-auto"><div class="flex justify-between items-center mb-4"><h4 class="text-lg font-bold">Editar Produto</h4><span onclick="closeModal()" class="cursor-pointer text-2xl">&times;</span></div><form id="modal-form" class="space-y-3"><input type="hidden" id="modal-rowIndex" name="rowIndex"><input type="hidden" id="modal-id" name="id"><div><label class="text-xs font-bold text-gray-500">Nome</label><input id="modal-nome" name="nome" class="w-full border p-2 rounded"></div><div class="flex gap-2"><div class="flex-1"><label class="text-xs font-bold text-gray-500">Preço</label><input id="modal-preco" name="preco" type="number" step="0.01" class="w-full border p-2 rounded"></div><div class="flex-1"><label class="text-xs font-bold text-gray-500">Categoria</label><input id="modal-categoria" name="categoria" class="w-full border p-2 rounded"></div></div><div><label class="text-xs font-bold text-gray-500">Descrição</label><textarea id="modal-descricao" name="descricao" rows="2" class="w-full border p-2 rounded"></textarea></div><div><label class="text-xs font-bold text-gray-500">Foto URL</label><input id="modal-foto_url" name="foto_url" class="w-full border p-2 rounded"></div><div><label class="text-xs font-bold text-gray-500">Status</label><select id="modal-ativo" name="ativo" class="w-full border p-2 rounded bg-white"><option value="SIM">Ativo</option><option value="NAO">Inativo</option></select></div><button id="modal-submit-btn" type="submit" class="w-full flex justify-center py-2 px-4 border border-transparent rounded-md shadow-sm text-sm font-medium text-white bg-indigo-600 hover:bg-indigo-700">Salvar Alterações</button></form><div id="modal-msg"></div></div></div>
      ${scriptCRUD}</body></html>`;
  }

  // >>> TELA: Gerenciar Insumos
  else if (templateName === 'GerenciarInsumos') {
    html = `<html>${head}<body class="p-0 bg-gray-50"><div id="loader" class="loader" style="margin-top: 2rem;"></div><div class="p-3 bg-white border-b sticky top-0 z-10 shadow-sm"><h3 class="font-bold text-lg text-gray-800 mb-2">Gerenciar Insumos</h3><input id="search" class="w-full border p-2 rounded bg-gray-50 focus:bg-white" placeholder="Buscar..."></div><div id="list" class="p-2 space-y-2"></div><div id="modal" class="modal"><div class="bg-white p-4 rounded w-96"><h4 class="font-bold mb-2">Editar Insumo</h4><form id="form" class="space-y-2"><input type="hidden" name="rowIndex" id="idx"><input type="hidden" name="id" id="eid"><input name="nome" id="enome" class="w-full border p-2 rounded"><input name="unidade" id="euni" class="w-full border p-2 rounded"><select name="ativo" id="eativo" class="w-full border p-2 rounded"><option value="SIM">Ativo</option><option value="NAO">Inativo</option></select><button id="btn-submit" class="w-full bg-blue-600 text-white p-2 rounded">Salvar</button></form><button onclick="closeModal()" class="mt-2 text-red-500 w-full">Cancelar</button></div></div>${getScriptForm('updateInsumo')}<script>function closeModal(){document.getElementById('modal').classList.remove('show')} function edit(id){const i=allData.find(x=>String(x.id)===String(id));if(!i)return;document.getElementById('idx').value=i.rowIndex;document.getElementById('eid').value=i.id;document.getElementById('enome').value=i.nome;document.getElementById('euni').value=i.unidade;document.getElementById('eativo').value=i.ativo;document.getElementById('modal').classList.add('show')} let allData=[];function load(){google.script.run.withSuccessHandler(d=>{allData=d;document.getElementById('loader').style.display='none';render(d)}).getInsumosParaAdmin()} function render(d){document.getElementById('list').innerHTML=d.map(i=>\`<div class="bg-white p-3 rounded border shadow-sm flex justify-between items-center"><div><div class="font-bold">\${i.nome}</div><div class="text-xs">Estoque: \${i.estoqueAtual} \${i.unidade}</div><div class="text-xs font-bold \${i.ativo==='SIM'?'text-green-600':'text-red-600'}">\${i.ativo}</div></div><button onclick="edit('\${i.id}')" class="text-indigo-600 border px-2 rounded font-bold">Editar</button></div>\`).join('')} document.getElementById('search').addEventListener('input',e=>{const t=e.target.value.toLowerCase();render(allData.filter(x=>x.nome.toLowerCase().includes(t)))});window.onload=load;</script></body></html>`;
  }

  // >>> TELA: Ficha Técnica
  else if (templateName === 'CadastroFichaTecnica') {
    html = `<html>${head}<body><h3 class="font-bold text-lg mb-4 text-gray-800">Ficha Técnica</h3><form id="main-form" class="space-y-3"><label class="block text-sm font-medium">1. Produto</label><select id="idItemCardapio" name="idItemCardapio" required class="w-full border p-2 rounded" onchange="loadReceita()"></select><div class="bg-gray-100 p-3 rounded text-sm border shadow-inner"><strong>Ingredientes:</strong><ul id="lista-receita" class="pl-4 list-disc mt-1">Carregando...</ul></div><label class="block text-sm font-medium mt-2">2. Insumo</label><select id="idInsumo" name="idInsumo" required class="w-full border p-2 rounded"></select><input name="quantidade" type="number" step="0.001" placeholder="Qtd Usada" required class="w-full border p-2 rounded"><button id="btn-submit" class="w-full bg-indigo-600 text-white p-2 rounded hover:bg-indigo-700">Adicionar</button></form><div id="loader" class="loader hidden mt-4"></div><div id="msg"></div>${getScriptForm('processarCadastroFichaTecnica')}<script>function loadData(){google.script.run.withSuccessHandler(d=>{const s1=document.getElementById('idItemCardapio'),s2=document.getElementById('idInsumo');s1.innerHTML='<option value="">Selecione...</option>'+d.cardapio.map(i=>\`<option value="\${i.id}">\${i.nome}</option>\`);s2.innerHTML='<option value="">Selecione...</option>'+d.insumos.map(i=>\`<option value="\${i.id}">\${i.nome}</option>\`);}).getDadosParaFormularios()} function loadReceita(){const id=document.getElementById('idItemCardapio').value;if(!id)return;document.getElementById('lista-receita').innerHTML='Loading...';google.script.run.withSuccessHandler(l=>{document.getElementById('lista-receita').innerHTML=l.map(i=>\`<li class="flex justify-between">\${i.quantidade} \${i.nomeInsumo} <span class="btn-delete" onclick="del('\${i.idFicha}','\${id}')">[x]</span></li>\`).join('')}).getFichaTecnicaDetalhada(id)} function del(f,p){if(confirm('Remover?'))google.script.run.withSuccessHandler(()=>loadReceita()).deleteIngredienteFichaTecnica(f,p)} window.onload=loadData;</script></body></html>`;
  }

  // >>> TELA: Gerenciar Adicionais
  else if (templateName === 'GerenciarAdicionais') {
    html = `<html>${head}<body><h3 class="font-bold text-lg mb-4 text-gray-800">Adicionais</h3><form id="main-form" class="space-y-3 bg-white p-3 border rounded"><input name="nome" placeholder="Nome" required class="w-full border p-2 rounded"><div class="flex gap-2"><input name="preco" type="number" step="0.01" placeholder="Preço" required class="w-full border p-2 rounded"><input name="categoria" placeholder="Categoria" required class="w-full border p-2 rounded"></div><button id="btn-submit" class="w-full bg-green-600 text-white p-2 rounded">Criar</button></form><div id="msg"></div><div id="list" class="space-y-2 mt-4"></div><div id="modal" class="modal"><div class="bg-white p-4 rounded w-96"><h4 class="font-bold mb-2">Editar</h4><form id="form-edit" class="space-y-2"><input type="hidden" id="eid"><input id="enome" class="w-full border p-2 rounded"><input id="epreco" type="number" step="0.01" class="w-full border p-2 rounded"><select id="eativo" class="w-full border p-2 rounded"><option value="SIM">Ativo</option><option value="NAO">Inativo</option></select><button class="w-full bg-blue-600 text-white p-2 rounded">Salvar</button></form><button onclick="closeModal()" class="mt-2 text-red-500 w-full">Cancelar</button></div></div>${getScriptForm('criarAdicional')}<script>function closeModal(){document.getElementById('modal').classList.remove('show')} function edit(id){const i=allData.find(x=>String(x.id)===String(id));document.getElementById('eid').value=i.id;document.getElementById('enome').value=i.nome;document.getElementById('epreco').value=i.preco;document.getElementById('eativo').value=i.ativo;document.getElementById('modal').classList.add('show')} let allData=[];function loadData(){google.script.run.withSuccessHandler(l=>{allData=l;document.getElementById('list').innerHTML=l.map(i=>\`<div class="flex justify-between border p-2 rounded bg-white"><span>\${i.nome} (R$ \${i.preco})</span><div class="flex gap-2"><span class="font-bold \${i.ativo==='SIM'?'text-green-600':'text-red-600'}">\${i.ativo}</span><button onclick="edit('\${i.id}')" class="text-blue-600 font-bold">Edit</button></div></div>\`).join('')}).getAdicionaisParaAdmin()} document.getElementById('form-edit').addEventListener('submit',e=>{e.preventDefault();const d={id:document.getElementById('eid').value,nome:document.getElementById('enome').value,preco:document.getElementById('epreco').value,categoria:'Geral',ativo:document.getElementById('eativo').value};google.script.run.withSuccessHandler(()=>{closeModal();loadData()}).updateAdicional(d)});window.onload=loadData;</script></body></html>`;
  }

  // >>> TELA: Gerenciar Motoqueiros (NOVO!)
  else if (templateName === 'GerenciarMotoqueiros') {
    html = `<html>${head}<body><h3 class="font-bold text-lg mb-4 text-gray-800">Gerenciar Motoqueiros</h3><form id="main-form" class="space-y-3 bg-white p-3 border rounded"><input name="nome" placeholder="Nome do Entregador" required class="w-full border p-2 rounded"><input name="telefone" placeholder="Telefone (WhatsApp)" class="w-full border p-2 rounded"><input name="placa" placeholder="Placa / Veículo" class="w-full border p-2 rounded"><button id="btn-submit" class="w-full bg-green-600 text-white p-2 rounded font-bold">Cadastrar</button></form><div id="msg"></div><div id="list" class="space-y-2 mt-4">Loading...</div>${getScriptForm('salvarMotoqueiro')}<script>function loadMotos(){google.script.run.withSuccessHandler(l=>{document.getElementById('list').innerHTML=l.length?l.map(m=>\`<div class="bg-white p-3 border rounded shadow-sm flex justify-between items-center"><div><div class="font-bold">\${m.nome}</div><div class="text-xs text-gray-500">\${m.telefone} | \${m.placa}</div></div><button onclick="del('\${m.id}')" class="btn-delete">X</button></div>\`).join(''):'<p class="text-gray-500">Nenhum motoqueiro.</p>'}).getMotoqueirosAdmin()} function del(id){if(confirm('Excluir?'))google.script.run.withSuccessHandler(loadMotos).excluirMotoqueiro(id)} window.onload=loadMotos;</script></body></html>`;
  }

  // >>> TELA: Venda Balcão (NOVO - PDV)
  else if (templateName === 'VendaBalcao') {
    html = `<html>${head}<body class="bg-gray-100 p-2">
       <div class="bg-white p-4 rounded shadow-lg max-w-md mx-auto">
         <h3 class="text-xl font-bold text-gray-800 mb-4 border-b pb-2">PDV - Balcão</h3>
         
         <!-- Cliente -->
         <div class="mb-4 grid grid-cols-2 gap-2">
            <input id="cli-nome" class="border p-2 rounded text-sm" placeholder="Nome Cliente (Opcional)">
            <select id="cli-pag" class="border p-2 rounded text-sm bg-white"></select>
         </div>

         <!-- Adicionar Item -->
         <div class="bg-gray-50 p-3 rounded border mb-4">
            <label class="text-xs font-bold text-gray-500">ADICIONAR PRODUTO</label>
            <select id="sel-prod" class="w-full border p-2 rounded mb-2 bg-white" onchange="updatePrice()"></select>
            <div class="flex gap-2">
               <input id="qtd" type="number" value="1" min="1" class="w-20 border p-2 rounded text-center">
               <button onclick="addItem()" class="flex-1 bg-blue-600 text-white font-bold rounded hover:bg-blue-700">+ Adicionar</button>
            </div>
         </div>

         <!-- Lista -->
         <div id="cart-list" class="space-y-2 max-h-60 overflow-y-auto mb-4 border-t pt-2">
            <p class="text-center text-gray-400 text-sm py-4">Carrinho vazio.</p>
         </div>

         <!-- Total -->
         <div class="flex justify-between items-center text-xl font-bold text-gray-800 border-t pt-4 mb-4">
            <span>Total:</span>
            <span id="total">R$ 0.00</span>
         </div>

         <button onclick="finalizar()" id="btn-fin" class="w-full bg-green-600 text-white py-3 rounded font-bold text-lg hover:bg-green-700 shadow">FINALIZAR VENDA</button>
       </div>

       <script>
          let cardapio = [];
          let carrinho = [];
          let pagamentos = [];

          function init() {
             // Carrega Cardápio
             google.script.run.withSuccessHandler(data => {
                cardapio = data;
                const sel = document.getElementById('sel-prod');
                sel.innerHTML = '<option value="">Selecione...</option>';
                data.forEach(p => {
                   const opt = document.createElement('option');
                   opt.value = p.id;
                   opt.textContent = p.nome + ' - R$ ' + p.preco.toFixed(2);
                   sel.appendChild(opt);
                });
             }).getCardapioParaAdmin();

             // Carrega Pagamentos
             google.script.run.withSuccessHandler(pags => {
                 const sel = document.getElementById('cli-pag');
                 pags.forEach(p => {
                    const opt = document.createElement('option');
                    opt.value = p.metodo;
                    opt.textContent = p.metodo;
                    sel.appendChild(opt);
                 });
             }).getPagamentosAdmin();
          }

          function addItem() {
             const id = document.getElementById('sel-prod').value;
             const qtd = parseInt(document.getElementById('qtd').value);
             if(!id || qtd < 1) return;
             
             const prod = cardapio.find(p => p.id === id);
             carrinho.push({ ...prod, qtde: qtd });
             renderCart();
          }

          function renderCart() {
             const list = document.getElementById('cart-list');
             list.innerHTML = '';
             let total = 0;
             
             carrinho.forEach((item, idx) => {
                total += item.preco * item.qtde;
                list.innerHTML += \`
                   <div class="flex justify-between items-center text-sm border-b pb-1">
                      <div>\${item.qtde}x \${item.nome}</div>
                      <div class="flex gap-2 items-center">
                         <span>R$ \${(item.preco * item.qtde).toFixed(2)}</span>
                         <button onclick="remItem(\${idx})" class="text-red-500 font-bold">x</button>
                      </div>
                   </div>\`;
             });
             
             document.getElementById('total').textContent = 'R$ ' + total.toFixed(2);
             if(carrinho.length === 0) list.innerHTML = '<p class="text-center text-gray-400 text-sm py-4">Carrinho vazio.</p>';
          }

          function remItem(idx) {
             carrinho.splice(idx, 1);
             renderCart();
          }

          function finalizar() {
             if(carrinho.length === 0) return alert('Carrinho vazio!');
             
             const btn = document.getElementById('btn-fin');
             btn.disabled = true;
             btn.textContent = 'Processando...';

             const pedido = {
                clienteNome: document.getElementById('cli-nome').value || 'Balcão',
                clienteTelefone: '',
                pagamento: document.getElementById('cli-pag').value || 'Dinheiro',
                itens: carrinho,
                bairro: 'Balcão',
                logradouro: 'Retirada',
                numero: '0'
             };

             google.script.run
                .withSuccessHandler(res => {
                   alert('Venda Realizada!');
                   carrinho = [];
                   renderCart();
                   document.getElementById('cli-nome').value = '';
                   btn.disabled = false;
                   btn.textContent = 'FINALIZAR VENDA';
                })
                .withFailureHandler(err => {
                   alert('Erro: ' + err.message);
                   btn.disabled = false;
                   btn.textContent = 'FINALIZAR VENDA';
                })
                .processarPedidoDelivery(JSON.stringify(pedido));
          }

          window.onload = init;
       </script>
    </body></html>`;
  }

  // >>> TELA: Ligar Adicionais
  else if (templateName === 'LigarAdicionais') {
    html = `<html>${head}<body><h3 class="font-bold mb-4">Vincular</h3><form id="main-form" class="space-y-3"><select id="prod" name="idProduto" class="w-full border p-2 rounded" onchange="loadL()"></select><div id="chk" class="max-h-60 overflow-y-auto bg-white border p-2 rounded p-2 text-sm"></div><button id="btn-submit" class="w-full bg-indigo-600 text-white p-2 rounded">Salvar</button></form><div id="msg"></div>${getScriptForm('processarLinkAdicionais')}<script>let adds=[];window.onload=()=>{google.script.run.withSuccessHandler(d=>{const s=document.getElementById('prod');s.innerHTML='<option>Selecione...</option>'+d.produtos.map(p=>\`<option value="\${p.id}">\${p.nome}</option>\`).join('');adds=d.adicionais}).getDadosParaLigarAdicionais()};function loadL(){const id=document.getElementById('prod').value;if(!id)return;document.getElementById('chk').innerHTML='Loading...';google.script.run.withSuccessHandler(ids=>{document.getElementById('chk').innerHTML=adds.map(a=>\`<label class="checkbox-label"><input type="checkbox" value="\${a.id}" class="chk-addon form-checkbox" \${ids.includes(a.id)?'checked':''}>\${a.nome} (+R$ \${a.preco})</label>\`).join('')}).getLinksAtuais(id)}</script></body></html>`;
  }

  // >>> TELA: Gerenciar Taxas
  else if (templateName === 'GerenciarTaxas') {
    html = `<html>${head}<body class="space-y-4"><h3 class="font-bold">Taxas Entrega</h3><form id="main-form" class="p-3 bg-white border rounded shadow-sm space-y-3"><div><label class="block text-sm">Bairro</label><input name="bairro" required class="w-full border p-2 rounded"></div><div><label class="block text-sm">Valor (R$)</label><input name="taxa" type="number" step="0.01" required class="w-full border p-2 rounded"></div><button id="btn-submit" class="w-full bg-green-600 text-white p-2 rounded">Salvar</button></form><div id="msg"></div><ul id="list" class="mt-4 space-y-2 text-sm"></ul>${getScriptForm('salvarTaxaEntrega')}<script>function loadTaxas(){google.script.run.withSuccessHandler(l=>{document.getElementById('list').innerHTML=l.map(i=>\`<li class="flex justify-between bg-white p-2 border rounded"><span>\${i.bairro}: R$ \${i.taxa.toFixed(2)}</span><button onclick="del('\${i.bairro}')" class="btn-delete">X</button></li>\`).join('')}).getTaxasParaAdmin()}function del(b){if(confirm('Excluir?'))google.script.run.withSuccessHandler(loadTaxas).excluirTaxaEntrega(b)}window.onload=loadTaxas</script></body></html>`;
  }

  // >>> TELA: Gerenciar Pagamentos
  else if (templateName === 'GerenciarPagamentos') {
    html = `<html>${head}<body class="space-y-4"><h3 class="font-bold">Pagamentos</h3><form id="main-form" class="flex gap-2"><input name="metodo" placeholder="Nome" required class="flex-1 border p-2 rounded"><button id="btn-submit" class="bg-green-600 text-white p-2 rounded">+</button></form><div id="msg"></div><ul id="list" class="mt-4 space-y-2 text-sm"></ul>${getScriptForm('salvarPagamento')}<script>function loadPags(){google.script.run.withSuccessHandler(l=>{document.getElementById('list').innerHTML=l.map(i=>\`<li class="flex justify-between bg-white p-2 border rounded"><span>\${i.metodo}</span><button onclick="del('\${i.metodo}')" class="btn-delete">X</button></li>\`).join('')}).getPagamentosAdmin()}function del(m){if(confirm('Excluir?'))google.script.run.withSuccessHandler(loadPags).excluirPagamento(m)}window.onload=loadPags</script></body></html>`;
  }

  // >>> TELA: Config WhatsApp
  else if (templateName === 'ConfigWhatsApp') {
    html = `<html>${head}<body><h3 class="font-bold text-green-700">WhatsApp</h3><form id="main-form" class="space-y-3"><input name="numero" id="numero" placeholder="5511999998888" required class="w-full border p-2 rounded"><button id="btn-submit" class="w-full bg-green-600 text-white p-2 rounded">Salvar</button></form><div id="msg"></div>${getScriptForm('salvarWhatsAppConfig')}<script>google.script.run.withSuccessHandler(n=>{if(n)document.getElementById('numero').value=n}).getWhatsAppConfigAtual()</script></body></html>`;
  }

  // >>> TELA: Registrar Compra Lote
  else if (templateName === 'RegistrarCompraLote') {
    html = `<html>${head}<body class="space-y-4"><div class="flex justify-between border-b pb-2"><h3 class="font-bold">Entrada Nota</h3><button onclick="add()" class="bg-blue-100 text-blue-700 px-2 rounded">+ Item</button></div><form id="main-form"><div id="rows" class="max-h-80 overflow-y-auto space-y-2"></div><div class="mt-2 border-t pt-2 font-bold flex justify-between"><span>Total:</span><span id="tot">R$ 0.00</span></div><button id="btn-submit" class="w-full bg-indigo-600 text-white p-2 rounded mt-2">Processar</button></form><div id="loader" class="loader hidden"></div><div id="msg"></div><script>const f=document.getElementById('main-form');let cache=[];function load(){google.script.run.withSuccessHandler(d=>{cache=d.insumos;add()}).getDadosParaFormularios()}function add(){const d=document.createElement('div');d.className='flex gap-1 items-end row';d.innerHTML=\`<div class="flex-1"><select class="w-full border p-1 sel"></select></div><div class="w-16"><input type="number" step="0.01" class="w-full border p-1 qtd" placeholder="Qtd"></div><div class="w-20"><input type="number" step="0.01" class="w-full border p-1 prc" placeholder="R$"></div><button type="button" onclick="this.parentNode.remove();upd()" class="text-red-500">X</button>\`;const s=d.querySelector('.sel');s.innerHTML='<option value="">Sel...</option>'+cache.map(i=>\`<option value="\${i.id}">\${i.nome}</option>\`).join('');d.querySelectorAll('input').forEach(i=>i.addEventListener('input',upd));document.getElementById('rows').appendChild(d)}function upd(){let t=0;document.querySelectorAll('.row').forEach(r=>{t+=(r.querySelector('.qtd').value*r.querySelector('.prc').value)||0});document.getElementById('tot').textContent='R$ '+t.toFixed(2)}f.addEventListener('submit',e=>{e.preventDefault();const i=[];document.querySelectorAll('.row').forEach(r=>i.push({idInsumo:r.querySelector('.sel').value,quantidade:r.querySelector('.qtd').value,preco:r.querySelector('.prc').value}));if(i.length==0)return;document.getElementById('loader').style.display='block';google.script.run.withSuccessHandler(r=>{document.getElementById('loader').style.display='none';document.getElementById('msg').textContent=r;document.getElementById('rows').innerHTML='';add();upd()}).processarCompraLote(i)});window.onload=load</script></body></html>`;
  }

  // >>> TELA: Relatórios
  else if (templateName === 'CentralRelatorios') {
    html = `<html>${head}<body class="space-y-3"><h3 class="font-bold">Relatórios</h3><div onclick="gen('ESTOQUE_BAIXO')" class="bg-white p-3 border rounded cursor-pointer hover:bg-orange-50 text-orange-600 font-bold">Estoque Crítico</div><div onclick="gen('VENDAS_DETALHADA')" class="bg-white p-3 border rounded cursor-pointer hover:bg-green-50 text-green-600 font-bold">Extrato Vendas</div><div id="loader" class="loader hidden"></div><div id="msg" class="text-center font-bold mt-2"></div><script>function gen(t){document.getElementById('loader').style.display='block';document.getElementById('msg').textContent='Gerando...';google.script.run.withSuccessHandler(r=>{document.getElementById('loader').style.display='none';document.getElementById('msg').textContent=r}).gerarRelatorio(t)}</script></body></html>`;
  }

  // >>> TELA: Ajuste Estoque
  else if (templateName === 'AjustarEstoque') {
    html = `<html>${head}<body><h3 class="font-bold mb-4">Ajuste Manual</h3><form id="main-form" class="space-y-3"><select id="idInsumo" name="idInsumo" required class="w-full border p-2 rounded"></select><input name="contagemReal" type="number" placeholder="Real" required class="w-full border p-2 rounded"><input name="motivo" placeholder="Motivo" class="w-full border p-2 rounded"><button id="btn-submit" class="w-full bg-orange-600 text-white p-2 rounded">Ajustar</button></form><div id="msg"></div>${getScriptForm('processarAjusteEstoque')}<script>google.script.run.withSuccessHandler(d=>{const s=document.getElementById('idInsumo');d.insumos.forEach(i=>s.innerHTML+=\`<option value="\${i.id}">\${i.nome}</option>\`)}).getDadosParaFormularios()</script></body></html>`;
  }

  else {
    html = `<html>${head}<body><h3>Erro 404: ${templateName} não encontrado.</h3></body></html>`;
  }

  return HtmlService.createHtmlOutput(html).setSandboxMode(HtmlService.SandboxMode.IFRAME);
}
