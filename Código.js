// ====================================================================
// 1. CONSTANTES E CONFIGURAÇÕES (CONFIG)
// ====================================================================

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
// NOVAS ABAS (Para o Roadmap)
const ABA_MOTOQUEIROS = 'Motoqueiros';
const ABA_DASHBOARD = 'DASHBOARD';


// ====================================================================
// 2. ACESSO A DADOS (MODELS) - Camada de Acesso à Planilha
// ====================================================================

function getPlanilha() {
  try {
    return SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {
    Logger.log(`[CRITICAL] Erro ao acessar planilha: ${e.message}`);
    throw new Error('Erro fatal: Não foi possível acessar o banco de dados.');
  }
}

function getSheet(ss, name) {
  try {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      Logger.log(`[INFO] Aba "${name}" não encontrada. Criando...`);
      sheet = ss.insertSheet(name);
    }
    return sheet;
  } catch (e) {
    throw new Error(`Erro ao acessar aba ${name}: ${e.message}`);
  }
}

function getSheetOrCreate(ss, name, headers) {
  try {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      if (headers && headers.length > 0) sheet.appendRow(headers);
    }
    return sheet;
  } catch (e) {
    throw new Error(`Erro ao criar aba ${name}: ${e.message}`);
  }
}

// ====================================================================
// 2. Configuração Inicial (Versão POC - Auto-Criação)
// ====================================================================

/**
 * @function configurarPlanilha
 * @description Cria abas e cabeçalhos automaticamente para o teste do zero.
 */
function configurarPlanilha() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Definição da estrutura completa do Banco de Dados
  const estrutura = [
    { nome: ABA_ESTOQUE, headers: ['ID_INSUMO', 'NOME', 'UNIDADE', 'ESTOQUE_ATUAL', 'ATIVO'] },
    { nome: ABA_CARDAPIO, headers: ['ID_ITEM', 'NOME', 'PRECO', 'CATEGORIA', 'CUSTO', 'ATIVO', 'TEMPO_PREPARO', 'DISPONIVEL_DELIVERY', 'PRECO_SUGERIDO', 'MARGEM_LUCRO', 'CUSTO_PERCENTUAL', 'LUCRO_BRUTO', 'DESCRICAO', 'FOTO_URL'] },
    // ALTERAÇÃO AQUI: Adicionado 'ID_MOTOQUEIRO' ao final
    { nome: ABA_VENDAS, headers: ['ID_VENDA', 'DT_HORA_PEDIDO', 'JSON_PEDIDO', 'VL_TOTAL_PEDIDO', 'FORMAPAGAMENTO_PEDIDO', 'CLIENTE_NOME', 'CLIENTE_TELEFONE', 'LOGRADOURO', 'NUMERO', 'COMPLEMENTO', 'BAIRRO', 'CIDADE', 'PONTO_REFERENCIA', 'TAXA_ENTREGA', 'STATUS_PEDIDO', 'OBSERVACOES', 'ID_MOTOQUEIRO'] },
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
    // NOVA ABA
    { nome: ABA_MOTOQUEIROS, headers: ['ID', 'NOME', 'TELEFONE', 'PLACA', 'ATIVO'] }
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
        Logger.log(`[SETUP] Cabeçalhos inseridos na aba existente '${obj.nome}'.`);
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
// 3. MENU DA PLANILHA
// ====================================================================

/**
 * @function onOpen
 * @description Menu Principal - Versão Final (Opção A: Compra em Lote Padrão)
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🍔 Gestor Delivery')
    .addSubMenu(ui.createMenu('➕ Cadastros (Criar)')
      .addItem('📦 Insumo (Estoque)', 'showSidebarCadastroInsumo')
      .addItem('🍔 Item Cardápio', 'showSidebarCadastroCardapio')
      .addItem('🥓 Adicional', 'showSidebarGerenciarAdicionais')
      .addItem('🏍️ Motoqueiro', 'showSidebarGerenciarMotoqueiros')
      .addSeparator()
      .addItem('🔗 Ligar Adicionais', 'showSidebarLigarAdicionais')
      .addItem('📝 Ficha Técnica', 'showSidebarCadastroFichaTecnica')
      .addSeparator()
      .addItem('📍 Taxas de Entrega', 'showSidebarGerenciarTaxas')
      .addItem('💳 Formas de Pagamento', 'showSidebarGerenciarPagamentos'))
    .addSeparator()
    .addSubMenu(ui.createMenu('✏️ Gerenciar (Editar)')
      .addItem('Cardápio', 'showSidebarGerenciarCardapio')
      .addItem('Insumos', 'showSidebarGerenciarInsumos'))
    .addSeparator()
    .addSubMenu(ui.createMenu('💰 Vendas & Estoque')
      .addItem('🖥️ Venda Balcão (PDV)', 'showSidebarVendaBalcao')
      .addItem('📄 Entrada de Nota (Lote)', 'showSidebarRegistrarCompraLote')
      .addItem('📉 Ajuste Manual', 'showSidebarAjustarEstoque')
      .addItem('🔄 Recalcular Custos', 'uiRecalcularCustoDosProdutos')
      .addItem('📊 Atualizar Dashboard', 'uiAtualizarDashboard')
      .addItem('📑 Central de Relatórios', 'showSidebarRelatorios'))
    .addSeparator()
    .addItem('🚀 ABRIR APP DELIVERY', 'showSidebarLinkDelivery') // [NOVO]
    .addItem('👨‍🍳 ABRIR KDS COZINHA', 'showSidebarLinkKDS')     // [NOVO]
    .addSeparator()
    .addItem('📱 Configurar WhatsApp', 'showSidebarConfigWhatsApp')
    .addItem('🚨 LIMPAR DADOS TESTE', 'adminLimparDadosDeTeste')
    .addToUi();
}
// --- WRAPPERS (Funções que abrem as janelas laterais) ---

// [NOVOS WRAPPERS NECESSÁRIOS]
function showSidebarGerenciarMotoqueiros() { openSidebar('GerenciarMotoqueiros', 'Gerenciar Motoqueiros'); }
function showSidebarVendaBalcao() { openSidebar('VendaBalcao', 'PDV - Balcão'); }

// [WRAPPERS EXISTENTES - Mantenha ou certifique-se que estão lá]
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

// Função Helper para abrir Sidebars
function openSidebar(templateName, title) {
  try {
    const html = getHtmlTemplate(templateName);
    if (!html) throw new Error(`Template "${templateName}" não encontrado.`);
    SpreadsheetApp.getUi().showSidebar(html.setTitle(title));
  } catch (error) {
    SpreadsheetApp.getUi().alert(`Erro ao abrir janela: ${error.message}`);
  }
}

// ====================================================================
// [NOVO] ACESSO RÁPIDO AOS APPS (Links Dinâmicos)
// ====================================================================

// ⚠️ COLE AQUI O SEU LINK QUE FUNCIONA (O que termina em ...iz/exec)
const URL_REAL = "https://script.google.com/macros/s/AKfycbymrpOtjNN3cD0seBOTpSCwSmy1vXrBkgTHUJa8cOblI0QN5YRR4ySeel-CBasLJ-iz/exec";

function showSidebarLinkDelivery() {
  // Usa a URL fixa para garantir que abra o link certo
  abrirSidebarLink('📱 App do Cliente', URL_REAL, 'Copie ou clique abaixo para abrir o cardápio:');
}

function showSidebarLinkKDS() {
  // Adiciona o parametro da cozinha na URL fixa
  const kdsUrl = URL_REAL + "?page=cozinha";
  abrirSidebarLink('👨‍🍳 Monitor KDS', kdsUrl, 'Link exclusivo para a tela da cozinha:');
}

function abrirSidebarLink(titulo, url, texto) {
  const html = `
    <html>
      <head>
        <base target="_top">
        <script src="https://cdn.tailwindcss.com"></script>
        <link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet">
      </head>
      <body class="bg-gray-50 p-4 flex flex-col items-center justify-center h-screen">
        <h3 class="text-xl font-bold text-gray-800 mb-2 text-center">${titulo}</h3>
        <p class="text-sm text-gray-500 mb-6 text-center">${texto}</p>
        
        <a href="${url}" target="_blank" class="bg-indigo-600 hover:bg-indigo-700 text-white font-bold py-3 px-6 rounded-full shadow-lg flex items-center gap-2 transform transition hover:scale-105 text-center">
          <span class="material-icons">open_in_new</span> ABRIR SISTEMA
        </a>

        <div class="mt-8 w-full">
           <label class="text-xs font-bold text-gray-400 uppercase">Link Direto:</label>
           <input value="${url}" class="w-full text-xs p-2 border rounded bg-white text-gray-600 select-all" readonly>
           <p class="text-[10px] text-gray-400 mt-1 text-center">Envie este link para seus clientes/funcionários.</p>
        </div>
      </body>
    </html>
  `;

  const output = HtmlService.createHtmlOutput(html).setTitle(titulo).setWidth(300);
  SpreadsheetApp.getUi().showSidebar(output);
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

// ====================================================================
// 5. BACKEND - PROCESSAMENTO
// ====================================================================

// --- INSUMOS ---
function processarCadastroInsumo(formData) {

  try {
    const aba = getSheet(getPlanilha(), ABA_ESTOQUE);
    const lastRow = aba.getLastRow();
    const novoId = `IN-${lastRow}`; // Simplificado para POC, ideal usar UUID ou Timestamp

    // Tratamento robusto de número (troca vírgula por ponto)
    let estoqueStr = String(formData.estoque || '0').replace(',', '.');
    const estoqueInicial = Number(formData.estoque) || 0;

    aba.appendRow([novoId, String(formData.nome).trim(), String(formData.unidade).trim(), estoqueInicial, 'SIM']);

    // REGRA AJUSTADA: Se tiver estoque, loga "Cadastro Inicial". 
    // Se for 0, loga "Item Criado (Sem Estoque)" para confirmar que a aba Log funciona.
    if (estoqueInicial > 0) {
      registrarLogEstoque(novoId, formData.nome, estoqueInicial, 'Cadastro Inicial');
    } else {
      registrarLogEstoque(novoId, formData.nome, 0, 'Item Criado (Estoque Zero)');
    }

    return 'Insumo cadastrado com sucesso!';
  } catch (e) {
    Logger.log(`[ERRO] processarCadastroInsumo: ${e.message}`);
    throw new Error(`Erro ao salvar: ${e.message}`);
  }
}

function getInsumosParaAdmin() {
  const data = getSheet(getPlanilha(), ABA_ESTOQUE).getDataRange().getValues();
  data.shift();
  return data.map((r, i) => ({
    rowIndex: i + 2, id: r[0], nome: r[1], unidade: r[2], estoqueAtual: Number(r[3]), ativo: r[4]
  })).filter(x => x.id);
}

function updateInsumo(formData) {
  const sheet = getSheet(getPlanilha(), ABA_ESTOQUE);
  const row = Number(formData.rowIndex);
  sheet.getRange(row, 2).setValue(formData.nome);
  sheet.getRange(row, 3).setValue(formData.unidade);
  sheet.getRange(row, 5).setValue(formData.ativo);
  return 'Insumo atualizado!';
}


// --- CARDÁPIO ---
function processarCadastroCardapio(formData) {
  const aba = getSheet(getPlanilha(), ABA_CARDAPIO);
  const lastRow = aba.getLastRow();
  const novoId = `CD-${lastRow}`;
  const preco = Number(formData.preco);

  // Colunas: ID, NOME, PRECO, CATEGORIA, CUSTO(0), ATIVO, ..., DESCRICAO, FOTO
  aba.appendRow([
    novoId,
    String(formData.nome).trim(),
    preco,
    String(formData.categoria).trim(),
    0,
    String(formData.ativo).toUpperCase(),
    '', '', '', '', '', '',
    String(formData.descricao || ''),
    String(formData.foto_url || '')
  ]);
  return 'Item salvo com sucesso!';
}

function processarCadastroFichaTecnica(formData) {
  const aba = getSheet(getPlanilha(), ABA_FICHA_TECNICA);
  const lastRow = aba.getLastRow();
  const novoId = `FT-${lastRow}`;

  // Dados para cache visual na planilha
  const nomeItem = getNomeItemPorID(formData.idItemCardapio) || 'N/A';
  const nomeInsumo = getNomeInsumoPorID(getSheet(getPlanilha(), ABA_ESTOQUE), formData.idInsumo) || 'N/A';

  aba.appendRow([
    novoId,
    formData.idItemCardapio,
    formData.idInsumo,
    Number(formData.quantidade),
    nomeItem,
    nomeInsumo
  ]);

  recalcularCustoDosProdutos(formData.idItemCardapio);
  return 'Ingrediente adicionado à Ficha Técnica!';
}

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

function getFichaTecnicaDetalhada(idProduto) {
  const ss = getPlanilha();
  const dados = getSheet(ss, ABA_FICHA_TECNICA).getDataRange().getValues();
  dados.shift();
  return dados
    .filter(r => r[1] === idProduto)
    .map(r => ({
      idFicha: r[0],
      idProduto: r[1],
      nomeInsumo: r[5] || 'Insumo',
      quantidade: Number(r[3])
    }));
}

function deleteIngredienteFichaTecnica(idFicha, idProduto) {
  const sheet = getSheet(getPlanilha(), ABA_FICHA_TECNICA);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === idFicha) {
      sheet.deleteRow(i + 1);
      recalcularCustoDosProdutos(idProduto);
      return 'Ingrediente removido.';
    }
  }
  throw new Error('Ingrediente não encontrado.');
}

// ====================================================================
// 7. ADMIN - GETTERS/UPDATES
// ====================================================================

function getCardapioParaAdmin() {
  const data = getSheet(getPlanilha(), ABA_CARDAPIO).getDataRange().getValues();
  data.shift();
  return data.map((r, i) => ({
    rowIndex: i + 2, id: r[0], nome: r[1], preco: Number(r[2]), categoria: r[3],
    custo: Number(r[4]), ativo: r[5], descricao: r[12], foto_url: r[13]
  })).filter(x => x.id);
}

function updateProdutoCardapio(formData) {
  const sheet = getSheet(getPlanilha(), ABA_CARDAPIO);
  const row = Number(formData.rowIndex);
  sheet.getRange(row, 2).setValue(formData.nome);
  sheet.getRange(row, 3).setValue(Number(formData.preco));
  sheet.getRange(row, 4).setValue(formData.categoria);
  sheet.getRange(row, 6).setValue(formData.ativo);
  sheet.getRange(row, 13).setValue(formData.descricao);
  sheet.getRange(row, 15).setValue(formData.foto_url);
  return 'Produto atualizado!';
}



// ====================================================================
// 8. ADICIONAIS (CRUD & LINK)
// ====================================================================

function criarAdicional(formData) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const sheet = getSheetOrCreate(getPlanilha(), ABA_ADICIONAIS, ["ID_ADICIONAL", "NOME", "PRECO", "CATEGORIA", "ATIVO"]);
    const lastRow = sheet.getLastRow();
    const id = `AD-${lastRow}`;
    sheet.appendRow([id, formData.nome, Number(formData.preco.replace(',', '.')), formData.categoria, 'SIM']);
    return 'Adicional criado!';
  } finally {
    lock.releaseLock();
  }
}

function getAdicionaisParaAdmin() {
  const sheet = getSheet(getPlanilha(), ABA_ADICIONAIS);
  const data = sheet.getDataRange().getValues();
  data.shift();
  return data.map((r, i) => ({
    rowIndex: i + 2, id: r[0], nome: r[1], preco: Number(r[2]), categoria: r[3], ativo: r[4]
  })).filter(x => x.id);
}

function updateAdicional(formData) {
  const sheet = getSheet(getPlanilha(), ABA_ADICIONAIS);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(formData.id)) {
      const r = i + 1;
      sheet.getRange(r, 2).setValue(formData.nome);
      sheet.getRange(r, 3).setValue(Number(formData.preco));
      sheet.getRange(r, 4).setValue(formData.categoria);
      sheet.getRange(r, 5).setValue(formData.ativo);
      return 'Adicional atualizado!';
    }
  }
  throw new Error('Adicional não encontrado.');
}

function inativarAdicional(formData) {
  // Reutiliza update mudando ativo para NAO
  formData.ativo = 'NAO';
  return updateAdicional(formData);
}

function getDadosParaLigarAdicionais() {
  const prods = getCardapioParaAdmin().filter(p => p.ativo === 'SIM');
  const adds = getAdicionaisParaAdmin().filter(a => a.ativo === 'SIM');
  return { produtos: prods, adicionais: adds };
}

function getLinksAtuais(idProduto) {
  const sheet = getSheetOrCreate(getPlanilha(), ABA_PRODUTO_ADICIONAIS_LINK, ['ID_PRODUTO', 'ID_ADICIONAL', 'ATIVO']);
  const data = sheet.getDataRange().getValues();
  data.shift();
  return data
    .filter(r => String(r[0]) === String(idProduto) && r[2] === 'SIM')
    .map(r => r[1]);
}

function processarLinkAdicionais(data) {
  const sheet = getSheetOrCreate(getPlanilha(), ABA_PRODUTO_ADICIONAIS_LINK, ['ID_PRODUTO', 'ID_ADICIONAL', 'ATIVO']);
  const idProd = data.idProduto;
  const novosAdds = data.adicionais || [];

  // Estratégia simples: Ler tudo, filtrar o que NÃO é deste produto, adicionar os novos
  const allData = sheet.getDataRange().getValues();
  const header = allData.shift();
  const outrosLinks = allData.filter(r => String(r[0]) !== String(idProd));

  const novasLinhas = novosAdds.map(idAdd => [idProd, idAdd, 'SIM']);
  const finalData = [header, ...outrosLinks, ...novasLinhas];

  sheet.clearContents();
  if (finalData.length > 0) {
    sheet.getRange(1, 1, finalData.length, 3).setValues(finalData);
  }
  return 'Vínculos salvos!';
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

function recalcularCustoDosProdutos(idProdutoAlvo) {
  const ss = getPlanilha();
  const abaCardapio = getSheet(ss, ABA_CARDAPIO);
  const abaFicha = getSheet(ss, ABA_FICHA_TECNICA);
  const abaCustos = getSheet(ss, ABA_CUSTO_INSUMOS); // Aba de custos

  // Mapa de custos unitários dos insumos
  const mapCustos = {};
  const dadosCustos = abaCustos.getDataRange().getValues();
  // Assume que na aba Custo_Insumos a Coluna 0 é Nome, Coluna 1 é ID, Coluna 5 é Custo Unitário
  // Se não tiver dados lá, usa 0
  for (let i = 1; i < dadosCustos.length; i++) {
    const id = dadosCustos[i][1];
    const custo = Number(dadosCustos[i][5]);
    if (id) mapCustos[id] = custo;
  }

  // Receitas
  const receitas = abaFicha.getDataRange().getValues();

  // Itens Cardapio
  const itens = abaCardapio.getDataRange().getValues();

  for (let i = 1; i < itens.length; i++) {
    const idProd = itens[i][0];
    if (idProdutoAlvo && idProd !== idProdutoAlvo) continue;

    // Soma custo dos ingredientes
    let custoTotal = 0;
    receitas.forEach(r => {
      if (r[1] === idProd) {
        const idInsumo = r[2];
        const qtd = Number(r[3]);
        const custoUnit = mapCustos[idInsumo] || 0;
        custoTotal += (qtd * custoUnit);
      }
    });

    // Atualiza Coluna 5 (Custo) -> índice 4
    abaCardapio.getRange(i + 1, 5).setValue(custoTotal);
  }
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
  Logger.log(">>> doGet: Acessando Web App.");

  // Roteamento Simples
  const page = e.parameter.page;

  if (page === 'cozinha') {
    return HtmlService.createTemplateFromFile('DeliveryApp_Cozinha')
      .evaluate()
      .setTitle('👨‍🍳 Monitor da Cozinha (KDS)')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  }

  // Padrão: App do Cliente
  else {
    return HtmlService.createTemplateFromFile('DeliveryApp')
      .evaluate()
      .setTitle('Delivery Esquina do Zezo')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
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
  // Lock para evitar duplicidade em pedidos simultâneos
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) { // Espera até 30s
    throw new Error("O sistema está ocupado processando outro pedido. Tente novamente em alguns segundos.");
  }

  try {
    const pedido = JSON.parse(jsonPedido);
    const ss = getPlanilha();
    const abaVendas = getSheet(ss, ABA_VENDAS);
    const idVenda = `VD-${Date.now()}`;

    // 1. Calcula Total dos Produtos e Baixa Estoque
    let totalProdutos = 0;

    // Tratamento seguro para lista de itens
    let itens = [];
    if (Array.isArray(pedido.itens)) itens = pedido.itens;

    itens.forEach(item => {
      // Baixa Estoque (Função existente mantida)
      darBaixaEstoqueFichaTecnica(ss, item.id, item.qtde);

      let itemTotal = parseFloat(item.precoBase || item.preco);
      if (item.adicionais && Array.isArray(item.adicionais)) {
        item.adicionais.forEach(a => itemTotal += parseFloat(a.preco));
      }
      totalProdutos += (itemTotal * item.qtde);
    });

    // 2. Calcula Taxa de Entrega
    const taxaEntrega = getTaxaEntregaPorBairro(pedido.bairro);

    // 3. Total Final
    const totalFinal = totalProdutos + taxaEntrega;

    // 4. Grava na Planilha (COM COLUNA DE MOTOQUEIRO VAZIA NO FINAL)
    abaVendas.appendRow([
      idVenda,
      new Date(),
      jsonPedido,
      totalFinal,
      pedido.pagamento,
      pedido.clienteNome,
      pedido.clienteTelefone,
      pedido.logradouro,
      pedido.numero,
      '', // Complemento
      pedido.bairro,
      '', // Cidade
      pedido.referencia,
      taxaEntrega,
      'NOVO',            // Coluna 15 (O)
      '',                // Coluna 16 (P) - Observações
      ''                 // Coluna 17 (Q) - ID_MOTOQUEIRO (NOVO)
    ]);

    // 5. Gera Mensagem WhatsApp
    const msg = gerarMensagemWhatsApp(ss, idVenda, pedido, totalProdutos, taxaEntrega, new Date());

    return { status: 'sucesso', idVenda, mensagemWhatsApp: msg };

  } catch (e) {
    Logger.log(`[VENDA][ERRO] ${e.message}`);
    throw new Error(`Erro ao processar venda: ${e.message}`);
  } finally {
    lock.releaseLock();
  }
}

function darBaixaEstoqueFichaTecnica(ss, idProd, qtdVenda) {
  const abaFicha = getSheet(ss, ABA_FICHA_TECNICA);
  const abaEstoque = getSheet(ss, ABA_ESTOQUE);
  const dataFicha = abaFicha.getDataRange().getValues();

  const ingredientes = dataFicha.filter(r => r[1] === idProd);
  ingredientes.forEach(ing => {
    const idInsumo = ing[2];
    const qtdReceita = Number(ing[3]);
    const baixaTotal = qtdReceita * qtdVenda;
    atualizarEstoque(abaEstoque, idInsumo, -baixaTotal);
    registrarLogEstoque(idInsumo, 'Baixa Venda', -baixaTotal, `Venda Prod ${idProd}`);
  });
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

  // [NOVO] 1. Busca Motoqueiros para traduzir ID -> Nome
  // Usa 'Motoqueiros' direto ou a constante ABA_MOTOQUEIROS se você definiu
  let mapMotos = {};
  try {
    const abaMotos = ss.getSheetByName('Motoqueiros');
    if (abaMotos) {
      const dadosMotos = abaMotos.getDataRange().getValues();
      // Cria um dicionário: { 'MOT-123': 'João', ... }
      dadosMotos.forEach(m => { if (m[0]) mapMotos[m[0]] = m[1]; });
    }
  } catch (e) { Logger.log("Erro ao ler motoqueiros: " + e.message); }

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

      // 3. Extração dos Itens (Mantido)
      let rawJson = null;
      let listaItens = [];

      try {
        rawJson = JSON.parse(row[2]);
      } catch (e) {
        rawJson = [];
      }

      if (rawJson) {
        if (Array.isArray(rawJson)) {
          listaItens = rawJson;
        } else if (rawJson.itens && Array.isArray(rawJson.itens)) {
          listaItens = rawJson.itens;
        }
      }

      // [NOVO] 4. Identificação do Motoqueiro
      // A coluna Q é o índice 16 (A=0 ... Q=16)
      const idMoto = row[16];
      const nomeMoto = mapMotos[idMoto] || '';

      return {
        rowIndex: index + 2,
        id: row[0],
        dataHora: dataPedido.toISOString(),
        itens: listaItens,
        total: row[3],
        pagamento: row[4],
        clienteNome: row[5],
        clienteTel: row[6],
        logradouro: row[7],
        numero: row[8],
        bairro: row[10],
        status: row[14],
        obs: row[15],
        // Novos campos para o Frontend ler:
        idMotoqueiro: idMoto,
        nomeMotoqueiro: nomeMoto
      };
    })
    .filter(p => p !== null)
    .sort((a, b) => new Date(a.dataHora) - new Date(b.dataHora));

  return pedidos;
}

/**
 * Atualiza o Status e executa lógicas de borda (Notificação ou Estorno).
 */
/**
 * Atualiza o status do pedido e vincula motoqueiro se necessário.
 * @param {string} idPedido - ID da Venda
 * @param {string} novoStatus - Novo Status
 * @param {string} idMotoqueiro - (Opcional) ID do Motoqueiro para vincular
 */
function atualizarStatusPedido(idPedido, novoStatus, idMotoqueiro) {
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

    // Busca o pedido pelo ID (Coluna A = índice 0)
    for (let i = 1; i < dados.length; i++) {
      if (String(dados[i][0]) === String(idPedido)) {
        linha = i + 1;
        jsonPedidoOriginal = dados[i][2];
        clienteNome = dados[i][5];
        clienteTel = dados[i][6];
        statusAtual = dados[i][14]; // Coluna O (índice 14)
        break;
      }
    }

    if (linha === -1) throw new Error(`Pedido ${idPedido} não encontrado.`);

    // Proteção: Não cancelar duas vezes
    if (statusAtual === 'CANCELADO' && novoStatus === 'CANCELADO') {
      return { success: true, message: 'Pedido já estava cancelado.' };
    }

    // 1. Grava o Novo Status na Coluna 15 (O)
    aba.getRange(linha, 15).setValue(novoStatus);

    // 2. LÓGICA DO MOTOQUEIRO (NOVO)
    // Se o status for SAIU_ENTREGA e um ID for fornecido, grava na Coluna 17 (Q)
    if (novoStatus === 'SAIU_ENTREGA' && idMotoqueiro) {
      // Coluna 17 corresponde à letra Q
      aba.getRange(linha, 17).setValue(idMotoqueiro);
      Logger.log(`[DELIVERY] Motoqueiro ${idMotoqueiro} vinculado ao pedido ${idPedido}`);
    }

    // 3. Lógica de Estorno (Se for cancelamento)
    if (novoStatus === 'CANCELADO') {
      estornarEstoquePedido(ss, jsonPedidoOriginal, idPedido);
    }

    SpreadsheetApp.flush();

    // 4. Lógica de Notificação WhatsApp
    let linkZap = null;
    if (clienteTel && (novoStatus === 'PREPARANDO' || novoStatus === 'SAIU_ENTREGA')) {
      const num = String(clienteTel).replace(/\D/g, '');
      let texto = '';

      if (novoStatus === 'PREPARANDO') {
        texto = `Olá ${clienteNome}! 👨‍🍳 Seu pedido *${idPedido}* começou a ser preparado.`;
      } else if (novoStatus === 'SAIU_ENTREGA') {
        texto = `🛵 Saiu para entrega! O pedido *${idPedido}* está a caminho.`;
        // Opcional: Se quiser incluir o nome do motoqueiro na mensagem, precisaria buscar o nome pelo ID aqui
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
// [NOVO] MÓDULO: GESTÃO DE MOTOQUEIROS (CRUD)
// ====================================================================

/**
 * Busca lista de motoqueiros ativos.
 */
function getMotoqueirosAdmin() {
  try {
    // Tenta acessar a aba, se não existir, cria (Fallback de segurança)
    const ss = getPlanilha();
    let sheet = ss.getSheetByName('Motoqueiros');

    if (!sheet) {
      sheet = ss.insertSheet('Motoqueiros');
      sheet.appendRow(['ID', 'NOME', 'TELEFONE', 'PLACA', 'ATIVO']);
      return [];
    }

    if (sheet.getLastRow() <= 1) return [];

    const data = sheet.getDataRange().getValues();
    data.shift(); // Remove cabeçalho

    return data.map((r, i) => ({
      id: r[0],
      nome: r[1],
      telefone: r[2],
      placa: r[3],
      ativo: r[4]
    })).filter(x => x.id && x.ativo === 'SIM');

  } catch (e) {
    Logger.log(`[MOTO][ERROR] ${e.message}`);
    return [];
  }
}

/**
 * Salva um novo motoqueiro.
 */
function salvarMotoqueiro(data) {
  try {
    if (!data.nome) throw new Error("O nome do entregador é obrigatório.");

    const ss = getPlanilha();
    let sheet = ss.getSheetByName('Motoqueiros');
    if (!sheet) {
      sheet = ss.insertSheet('Motoqueiros');
      sheet.appendRow(['ID', 'NOME', 'TELEFONE', 'PLACA', 'ATIVO']);
    }

    const id = `MOT-${Date.now()}`;
    sheet.appendRow([
      id,
      String(data.nome).toUpperCase().trim(),
      String(data.telefone || ''),
      String(data.placa || '').toUpperCase(),
      'SIM'
    ]);

    return "Motoqueiro cadastrado com sucesso!";
  } catch (e) {
    throw new Error(`Erro ao salvar: ${e.message}`);
  }
}

/**
 * Remove (exclusão física) um motoqueiro.
 */
function excluirMotoqueiro(id) {
  try {
    const sheet = getSheet(getPlanilha(), 'Motoqueiros');
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.deleteRow(i + 1);
        return "Motoqueiro removido.";
      }
    }
    throw new Error("Motoqueiro não encontrado.");
  } catch (e) {
    throw new Error(e.message);
  }
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
  let qtdPedidosTotal = 0; // Começa com 0 e conta apenas os válidos
  let mapPagamentos = {};
  let mapProdutos = {};

  dados.forEach(row => {
    // --- CORREÇÃO AQUI ---
    // A verificação deve ser feita DENTRO do loop, linha por linha
    const statusPedido = row[14]; // Coluna O
    if (statusPedido === 'CANCELADO') return; // Pula este pedido se estiver cancelado

    // Se não for cancelado, conta para o total
    qtdPedidosTotal++;

    // Mapeamento das colunas
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


/*
/* Funções Uteis
*/

/**
 * Converte um link de compartilhamento do Google Drive em um link direto para imagem.
 * Uso na planilha: =GERAR_LINK_IMAGEM(A2)
 *
 * @param {string} url O link de compartilhamento do Google Drive.
 * @return O link direto para uso no App.
 * @customfunction
 */
function GERAR_LINK_IMAGEM(url) {
  if (!url || url === "") return "";

  try {
    // Extrai o ID do arquivo
    var id = "";
    var parts = url.split("/");

    // Tenta achar o ID padrão
    for (var i = 0; i < parts.length; i++) {
      if (parts[i] === "d") {
        id = parts[i + 1];
        break;
      }
    }

    // Se não achou pelo padrão /d/, tenta pegar do parâmetro ?id=
    if (id === "") {
      var match = url.match(/id=([a-zA-Z0-9_-]+)/);
      if (match && match[1]) {
        id = match[1];
      }
    }

    if (id === "") return "Link Inválido";

    // Retorna o link de exportação direta
    return "https://lh3.googleusercontent.com/d/" + id;

  } catch (e) {
    return "Erro ao converter";
  }
}


// Limpa os testes

function adminLimparDadosDeTeste() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('ZERAR TUDO?', 'Isso apaga Vendas, Logs e Compras.', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  const ss = getPlanilha();
  [ABA_VENDAS, ABA_LOG_ESTOQUE, ABA_COMPRAS].forEach(n => {
    const s = ss.getSheetByName(n);
    if (s && s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, s.getLastColumn()).clearContent();
  });
  ui.alert('Limpeza concluída.');
}

// Recalcula os custos
function uiRecalcularCustoDosProdutos() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('Confirmar', 'Deseja recalcular os custos agora?', ui.ButtonSet.OK_CANCEL);
  if (response === ui.Button.OK) {
    recalcularCustoDosProdutos(null);
    ui.alert('✅ Custos recalculados com sucesso!');
  }
}

// Gera relatórios
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
// 11. HTML TEMPLATES (FÁBRICA - VERSÃO CORRIGIDA E UNIFICADA)
// ====================================================================

function getHtmlTemplate(templateName) {
  let html = '';

  // --- 1. HEAD E ESTILOS GERAIS (Componente Reutilizável) ---
  const head = `
    <head>
      <base target="_top">
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <script src="https://cdn.tailwindcss.com"></script>
      <style>
        body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; padding: 1rem; background-color: #f9fafb; }
        .loader { border: 3px solid #f3f3f3; border-top: 3px solid #4f46e5; border-radius: 50%; width: 24px; height: 24px; animation: spin 1s linear infinite; margin: 0 auto; }
        @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
        
        /* Estilos Específicos */
        .checkbox-label { display: flex; align-items: center; padding: 0.5rem; border: 1px solid #e5e7eb; border-radius: 0.375rem; cursor: pointer; transition: background 0.2s; margin-bottom: 0.5rem; }
        .checkbox-label:hover { background-color: #f3f4f6; }
        .form-checkbox { height: 1.25rem; width: 1.25rem; margin-right: 0.5rem; color: #4f46e5; border-radius: 4px; }
        
        /* Botão de Excluir */
        .btn-delete { color: #ef4444; font-weight: bold; cursor: pointer; padding: 2px 6px; border-radius: 4px; }
        .btn-delete:hover { background-color: #fee2e2; }
      </style>
    </head>
  `;

  // --- 2. GERADOR DE SCRIPT (Lógica de Envio Unificada) ---
  // Esta função cria o <script> que envia os dados para o Google Apps Script
  const getScriptForm = (funcName) => `
    <script>
      // Funções de UI (Loading e Mensagens)
      function showLoading() { 
        const l = document.getElementById('loader'); if(l) l.style.display = 'block';
        const b = document.getElementById('btn-submit'); if(b) b.disabled = true;
        const m = document.getElementById('msg') || document.getElementById('message'); if(m) m.textContent = '';
      }
      
      function hideLoading() { 
        const l = document.getElementById('loader'); if(l) l.style.display = 'none';
        const b = document.getElementById('btn-submit'); if(b) b.disabled = false;
      }
      
      function showMsg(txt, err=false) {
        const el = document.getElementById('msg') || document.getElementById('message');
        if(el) {
            el.textContent = txt;
            el.className = err ? 'text-red-600 font-bold mt-2 text-center' : 'text-green-600 font-bold mt-2 text-center';
        }
      }
      
      // Handler Principal de Envio
      const form = document.getElementById('main-form');
      if(form) {
          form.addEventListener('submit', function(e) {
            e.preventDefault();
            showLoading();
            
            const formData = {};
            new FormData(form).forEach((v, k) => formData[k] = v);

            // Lógica Especial: Checkboxes (Ligar Adicionais)
            if('${funcName}' === 'processarLinkAdicionais') {
               const checks = document.querySelectorAll('.chk-addon:checked');
               formData.adicionais = Array.from(checks).map(c => c.value);
            }

            google.script.run
              .withSuccessHandler(res => {
                 hideLoading();
                 showMsg(res);
                 
                 // Reseta o form (exceto na tela de vínculos para não perder a seleção)
                 if('${funcName}' !== 'processarLinkAdicionais') form.reset();
                 
                 // Gatilhos de Recarga (Hooks)
                 if(typeof loadData === 'function') loadData(); // Insumos/Adicionais
                 if(typeof loadTaxas === 'function') loadTaxas(); // Taxas
                 if(typeof carregarFichaDoItem === 'function') carregarFichaDoItem(); // Ficha Técnica
              })
              .withFailureHandler(err => {
                 hideLoading();
                 showMsg(err.message, true);
              })
              .${funcName}(formData);
          });
      }
    </script>
  `;

  // --- 3. TEMPLATES HTML ---

  // >>> Template: Cadastro Insumo
  if (templateName === 'CadastroInsumo') {
    html = `<html>${head}<body>
      <h3 class="font-bold text-lg mb-4 text-gray-800">Novo Insumo</h3>
      <form id="main-form" class="space-y-3">
        <input name="nome" placeholder="Nome do Insumo" required class="w-full border p-2 rounded">
        <input name="unidade" placeholder="Unidade (KG, UN)" required class="w-full border p-2 rounded">
        <input name="estoque" type="number" step="0.01" placeholder="Estoque Inicial" class="w-full border p-2 rounded">
        <button id="btn-submit" class="w-full bg-indigo-600 text-white p-2 rounded hover:bg-indigo-700">Salvar</button>
      </form>
      <div id="loader" class="loader hidden mt-4"></div>
      <div id="msg"></div>
      ${getScriptForm('processarCadastroInsumo')}
    </body></html>`;
  }

  // >>> Template: Gerenciar Pagamentos
  else if (templateName === 'GerenciarPagamentos') {
    html = `<html>${head}<body class="space-y-4">
        <h3 class="text-lg font-semibold text-gray-800">Formas de Pagamento</h3>
        <p class="text-sm text-gray-600">Defina o que aparece no checkout do cliente.</p>
        
        <form id="main-form" class="flex gap-2">
           <input name="metodo" required placeholder="Ex: Vale Refeição" class="flex-1 border p-2 rounded">
           <button id="btn-submit" class="bg-green-600 text-white p-2 rounded font-bold">+</button>
        </form>
        <div id="msg"></div>
        <div id="loader" class="loader hidden"></div>

        <ul id="lista-pag" class="space-y-2 text-sm mt-4 border-t pt-4">Carregando...</ul>

        ${getScriptForm('salvarPagamento')}
        <script>
           function loadPags() {
              google.script.run.withSuccessHandler(renderList).getPagamentosAdmin();
           }
           function renderList(lista) {
              const ul = document.getElementById('lista-pag');
              ul.innerHTML = '';
              if(lista.length === 0) ul.innerHTML = '<li class="text-gray-500 italic">Nenhum método cadastrado.</li>';
              
              lista.forEach(item => {
                 const li = document.createElement('li');
                 li.className = 'flex justify-between items-center bg-white p-3 rounded border shadow-sm';
                 li.innerHTML = \`
                    <span class="font-medium">\${item.metodo}</span>
                    <button onclick="excluir('\${item.metodo}')" class="text-red-500 hover:bg-red-50 p-1 rounded"><span class="material-icons text-base">delete</span></button>
                 \`;
                 ul.appendChild(li);
              });
           }
           function excluir(metodo) {
              if(!confirm('Remover ' + metodo + '?')) return;
              google.script.run.withSuccessHandler(() => { loadPags(); }).excluirPagamento(metodo);
           }
           window.onload = loadPags;
           // Hook após salvar
           const formEl = document.getElementById('main-form');
           formEl.addEventListener('submit', () => setTimeout(loadPags, 1000));
        </script>
      </body></html>`;
  }

  // >>> Template: Cadastro Cardapio
  else if (templateName === 'CadastroCardapio') {
    html = `<html>${head}<body>
      <h3 class="font-bold text-lg mb-4 text-gray-800">Novo Item Cardápio</h3>
      <form id="main-form" class="space-y-3">
        <input name="nome" placeholder="Nome do Item" required class="w-full border p-2 rounded">
        <input name="categoria" placeholder="Categoria (Lanches, Bebidas)" required class="w-full border p-2 rounded">
        <input name="preco" type="number" step="0.01" placeholder="Preço Venda (R$)" required class="w-full border p-2 rounded">
        <textarea name="descricao" placeholder="Descrição" class="w-full border p-2 rounded"></textarea>
        <input name="foto_url" placeholder="URL da Foto" class="w-full border p-2 rounded">
        <select name="ativo" class="w-full border p-2 rounded"><option value="SIM">Ativo</option><option value="NAO">Inativo</option></select>
        <button id="btn-submit" class="w-full bg-indigo-600 text-white p-2 rounded hover:bg-indigo-700">Salvar</button>
      </form>
      <div id="loader" class="loader hidden mt-4"></div>
      <div id="msg"></div>
      ${getScriptForm('processarCadastroCardapio')}
    </body></html>`;
  }

  // >>> Template: Ficha Técnica
  else if (templateName === 'CadastroFichaTecnica') {
    html = `<html>${head}<body>
      <h3 class="font-bold text-lg mb-4 text-gray-800">Ficha Técnica</h3>
      <form id="main-form" class="space-y-3">
        <label class="block text-sm font-medium text-gray-700">1. Item do Cardápio</label>
        <select id="idItemCardapio" name="idItemCardapio" required class="w-full border p-2 rounded" onchange="loadReceita()"></select>
        
        <div class="bg-gray-100 p-3 rounded text-sm border shadow-inner">
           <strong>Ingredientes Atuais:</strong>
           <ul id="lista-receita" class="pl-4 list-disc mt-1 space-y-1 text-gray-600">Carregando...</ul>
        </div>

        <label class="block text-sm font-medium text-gray-700 mt-2">2. Adicionar Insumo</label>
        <select id="idInsumo" name="idInsumo" required class="w-full border p-2 rounded"></select>
        <input name="quantidade" type="number" step="0.001" placeholder="Qtd Usada" required class="w-full border p-2 rounded">
        
        <button id="btn-submit" class="w-full bg-indigo-600 text-white p-2 rounded hover:bg-indigo-700">Adicionar</button>
      </form>
      <div id="loader" class="loader hidden mt-4"></div>
      <div id="msg"></div>
      
      ${getScriptForm('processarCadastroFichaTecnica')}
      <script>
        function loadData() {
           google.script.run.withSuccessHandler(data => {
              const s1 = document.getElementById('idItemCardapio');
              const s2 = document.getElementById('idInsumo');
              s1.innerHTML = '<option value="">Selecione...</option>';
              s2.innerHTML = '<option value="">Selecione...</option>';
              data.cardapio.forEach(i => s1.innerHTML += \`<option value="\${i.id}">\${i.nome}</option>\`);
              data.insumos.forEach(i => s2.innerHTML += \`<option value="\${i.id}">\${i.nome}</option>\`);
           }).getDadosParaFormularios();
        }
        function loadReceita() {
           const id = document.getElementById('idItemCardapio').value;
           if(!id) return;
           document.getElementById('lista-receita').innerHTML = '<li>Carregando...</li>';
           google.script.run.withSuccessHandler(ing => {
              const ul = document.getElementById('lista-receita');
              ul.innerHTML = '';
              if(ing.length === 0) ul.innerHTML = '<li>Nenhum ingrediente.</li>';
              ing.forEach(i => ul.innerHTML += \`<li class="flex justify-between">\${i.quantidade} \${i.nomeInsumo} <span class="btn-delete" onclick="remove('\${i.idFicha}','\${id}')">[x]</span></li>\`);
           }).getFichaTecnicaDetalhada(id);
        }
        function remove(idFicha, idProd) {
           if(!confirm('Remover ingrediente?')) return;
           google.script.run.withSuccessHandler(() => loadReceita()).deleteIngredienteFichaTecnica(idFicha, idProd);
        }
        window.onload = loadData;
      </script>
    </body></html>`;
  }

  // >>> Template: Registrar Compra em Lote (NOVO)
  else if (templateName === 'RegistrarCompraLote') {
    html = `<html>${head}<body class="space-y-4">
        <div class="flex justify-between items-center border-b pb-2">
           <h3 class="text-lg font-semibold text-gray-800">Entrada de Nota</h3>
           <button onclick="addRow()" class="bg-blue-100 text-blue-700 px-3 py-1 rounded text-sm font-bold hover:bg-blue-200">+ Item</button>
        </div>
        
        <p class="text-xs text-gray-500">Adicione os itens da nota fiscal e salve tudo de uma vez.</p>

        <form id="main-form" class="space-y-3">
           <div id="lista-itens" class="space-y-3 max-h-80 overflow-y-auto p-1">
              </div>
           
           <div class="pt-4 border-t">
              <div class="flex justify-between text-sm font-bold mb-2">
                 <span>Total da Nota:</span>
                 <span id="total-nota">R$ 0.00</span>
              </div>
              <button id="submit-btn" class="w-full bg-indigo-600 text-white p-3 rounded font-bold hover:bg-indigo-700">
                 Processar Entrada
              </button>
           </div>
        </form>
        
        <div id="loader" class="loader hidden mt-4"></div>
        <div id="msg"></div>

        <template id="row-template">
           <div class="flex gap-2 items-end bg-white p-2 rounded border shadow-sm item-row">
              <div class="flex-1">
                 <label class="text-[10px] text-gray-500 font-bold">INSUMO</label>
                 <select name="idInsumo" class="w-full border p-1 rounded text-sm select-insumo"></select>
              </div>
              <div class="w-16">
                 <label class="text-[10px] text-gray-500 font-bold">QTD</label>
                 <input type="number" step="0.01" name="qtd" class="w-full border p-1 rounded text-sm input-qtd" placeholder="0">
              </div>
              <div class="w-20">
                 <label class="text-[10px] text-gray-500 font-bold">R$ UN.</label>
                 <input type="number" step="0.01" name="preco" class="w-full border p-1 rounded text-sm input-preco" placeholder="0.00">
              </div>
              <button type="button" onclick="removeRow(this)" class="text-red-400 hover:text-red-600 mb-1">
                 <span class="material-icons text-lg">delete</span>
              </button>
           </div>
        </template>

        <script>
           // Configuração do envio manual (sobrescreve o padrão)
           const form = document.getElementById('main-form');
           
           form.addEventListener('submit', (e) => {
              e.preventDefault();
              const rows = document.querySelectorAll('.item-row');
              if(rows.length === 0) { alert('Adicione pelo menos um item.'); return; }
              
              const itens = [];
              let erro = false;
              
              rows.forEach(row => {
                 const id = row.querySelector('.select-insumo').value;
                 const qtd = row.querySelector('.input-qtd').value;
                 const preco = row.querySelector('.input-preco').value;
                 
                 if(!id || !qtd || !preco) erro = true;
                 
                 itens.push({ idInsumo: id, quantidade: qtd, preco: preco });
              });
              
              if(erro) { alert('Preencha todos os campos de todos os itens.'); return; }
              
              // Envia
              showLoading();
              google.script.run
                 .withSuccessHandler(res => {
                    hideLoading();
                    showMsg(res);
                    document.getElementById('lista-itens').innerHTML = ''; // Limpa
                    addRow(); // Adiciona uma nova limpa
                    updateTotal();
                 })
                 .withFailureHandler(err => {
                    hideLoading();
                    showMsg(err.message, true);
                 })
                 .processarCompraLote(itens);
           });

           // Lógica Dinâmica
           let insumosCache = [];

           function loadInsumos() {
              google.script.run.withSuccessHandler(d => {
                 insumosCache = d.insumos;
                 addRow(); // Adiciona a primeira linha ao carregar
              }).getDadosParaFormularios();
           }

           function addRow() {
              const template = document.getElementById('row-template');
              const clone = template.content.cloneNode(true);
              const select = clone.querySelector('.select-insumo');
              
              // Popula Select
              select.innerHTML = '<option value="">Selecione...</option>';
              insumosCache.forEach(i => {
                 select.innerHTML += \`<option value="\${i.id}">\${i.nome}</option>\`;
              });
              
              // Listeners para total
              const inputs = clone.querySelectorAll('input');
              inputs.forEach(i => i.addEventListener('input', updateTotal));

              document.getElementById('lista-itens').appendChild(clone);
           }

           function removeRow(btn) {
              btn.closest('.item-row').remove();
              updateTotal();
           }
           
           function updateTotal() {
              let total = 0;
              document.querySelectorAll('.item-row').forEach(row => {
                 const q = parseFloat(row.querySelector('.input-qtd').value) || 0;
                 const p = parseFloat(row.querySelector('.input-preco').value) || 0;
                 total += (q * p);
              });
              document.getElementById('total-nota').textContent = 'R$ ' + total.toFixed(2);
           }

           // Funções UI Padrão (Copiadas do seu head para garantir escopo local se necessário)
           function showLoading() { 
             document.getElementById('loader').style.display = 'block';
             document.getElementById('submit-btn').disabled = true;
             document.getElementById('msg').textContent = '';
           }
           function hideLoading() { 
             document.getElementById('loader').style.display = 'none';
             document.getElementById('submit-btn').disabled = false;
           }
           function showMsg(txt, err=false) {
             const el = document.getElementById('msg');
             el.textContent = txt;
             el.className = err ? 'text-red-600 font-bold mt-2 text-center' : 'text-green-600 font-bold mt-2 text-center';
           }

           window.onload = loadInsumos;
        </script>
    </body></html>`;
  }

  // >>> Template: Ajustar Estoque
  else if (templateName === 'AjustarEstoque') {
    html = `<html>${head}<body>
      <h3 class="font-bold text-lg mb-4 text-gray-800">Ajuste Manual</h3>
      <form id="main-form" class="space-y-3">
        <select id="idInsumo" name="idInsumo" required class="w-full border p-2 rounded"></select>
        <input name="contagemReal" type="number" step="0.01" placeholder="Contagem Real" required class="w-full border p-2 rounded">
        <input name="motivo" placeholder="Motivo (Perda, Quebra)" class="w-full border p-2 rounded">
        <button id="btn-submit" class="w-full bg-orange-600 text-white p-2 rounded hover:bg-orange-700">Ajustar</button>
      </form>
      <div id="loader" class="loader hidden mt-4"></div>
      <div id="msg"></div>
      ${getScriptForm('processarAjusteEstoque')}
      <script>
        google.script.run.withSuccessHandler(d => {
           const s = document.getElementById('idInsumo');
           d.insumos.forEach(i => s.innerHTML += \`<option value="\${i.id}">\${i.nome}</option>\`);
        }).getDadosParaFormularios();
      </script>
    </body></html>`;
  }

  // >>> Template: Central Relatórios (CORRIGIDO)
  else if (templateName === 'CentralRelatorios') {
    html = `<html>${head}<body>
      <h3 class="font-bold text-lg mb-4 text-gray-800">Relatórios Gerenciais</h3>
      
      <div class="space-y-3">
        <div onclick="gen('ESTOQUE_BAIXO')" class="bg-white p-4 border rounded shadow-sm cursor-pointer hover:bg-orange-50 transition flex items-center justify-between group">
           <span class="font-bold text-orange-700 group-hover:text-orange-800">📉 Estoque Crítico</span>
           <span class="material-icons text-orange-400">warning</span>
        </div>
        
        <div onclick="gen('VENDAS_DETALHADA')" class="bg-white p-4 border rounded shadow-sm cursor-pointer hover:bg-green-50 transition flex items-center justify-between group">
           <span class="font-bold text-green-700 group-hover:text-green-800">💰 Extrato de Vendas</span>
           <span class="material-icons text-green-400">receipt_long</span>
        </div>
      </div>

      <div id="loader" class="loader hidden mt-6"></div>
      <div id="msg" class="mt-4 text-center font-bold text-sm"></div>

      <script>
        function gen(tipo) {
           const loader = document.getElementById('loader');
           const msg = document.getElementById('msg');
           
           loader.style.display = 'block';
           msg.textContent = 'Gerando relatório... Aguarde.';
           msg.className = 'text-gray-500 mt-4 text-center font-bold text-sm';

           google.script.run
             .withSuccessHandler(res => {
                loader.style.display = 'none';
                msg.textContent = res;
                msg.className = 'text-green-600 mt-4 text-center font-bold text-sm bg-green-100 p-2 rounded border border-green-200';
             })
             .withFailureHandler(err => {
                loader.style.display = 'none';
                msg.textContent = 'Erro: ' + err.message;
                msg.className = 'text-red-600 mt-4 text-center font-bold text-sm bg-red-100 p-2 rounded border border-red-200';
             })
             .gerarRelatorio(tipo);
        }
      </script>
    </body></html>`;
  }

  // >>> Template: Gerenciar Adicionais
  else if (templateName === 'GerenciarAdicionais') {
    html = `<html>${head}<body>
      <h3 class="font-bold text-lg mb-4 text-gray-800">Adicionais</h3>
      
      <form id="main-form" class="space-y-3 border-b pb-4 mb-4 bg-white p-3 rounded shadow-sm">
        <h4 class="font-semibold text-sm text-gray-700">Novo Adicional</h4>
        <input name="nome" placeholder="Nome (Ex: Bacon)" required class="w-full border p-2 rounded">
        <div class="flex gap-2">
           <input name="preco" type="number" step="0.01" placeholder="Preço" required class="w-full border p-2 rounded">
           <input name="categoria" placeholder="Categoria" required class="w-full border p-2 rounded">
        </div>
        <button id="btn-submit" class="w-full bg-green-600 text-white p-2 rounded hover:bg-green-700">Criar</button>
      </form>
      
      <div id="loader" class="loader hidden"></div>
      <div id="msg"></div>

      <div class="mt-4">
         <h4 class="font-semibold text-sm mb-2 text-gray-700">Lista Atual</h4>
         <div id="lista" class="text-sm space-y-2">Carregando...</div>
      </div>

      ${getScriptForm('criarAdicional')}
      <script>
         function loadData() {
            google.script.run.withSuccessHandler(lista => {
               const div = document.getElementById('lista');
               div.innerHTML = '';
               if(lista.length === 0) div.innerHTML = '<p class="text-gray-500">Nenhum adicional.</p>';
               lista.forEach(i => {
                  div.innerHTML += \`<div class="flex justify-between border p-2 rounded bg-white">
                    <span>\${i.nome} (R$ \${i.preco})</span>
                    <span class="\${i.ativo==='SIM'?'text-green-600':'text-red-600'} font-bold">\${i.ativo}</span>
                  </div>\`;
               });
            }).getAdicionaisParaAdmin();
         }
         window.onload = loadData;
      </script>
    </body></html>`;
  }

  // >>> Template: Configurar WhatsApp
  else if (templateName === 'ConfigWhatsApp') {
    html = `<html>${head}<body>
      <h3 class="font-bold text-lg mb-4 text-green-700 flex items-center gap-2">
        <span>📱</span> Configurar WhatsApp
      </h3>
      <p class="text-sm text-gray-600 mb-4">Este é o número que receberá os pedidos.</p>
      
      <form id="main-form" class="space-y-4 bg-white p-4 rounded border shadow-sm">
        <div>
          <label class="block text-sm font-medium text-gray-700 mb-1">Número (55 + DDD + Número)</label>
          <input name="numero" id="numero" type="text" placeholder="5511999998888" required 
                 class="w-full border border-gray-300 p-3 rounded text-lg font-mono">
        </div>
        <button id="btn-submit" class="w-full bg-green-600 hover:bg-green-700 text-white font-bold p-3 rounded">Salvar</button>
      </form>
      <div id="loader" class="loader hidden mt-4"></div>
      <div id="msg"></div>
      
      ${getScriptForm('salvarWhatsAppConfig')}
      <script>
        function loadNum() {
           document.getElementById('loader').style.display = 'block';
           google.script.run.withSuccessHandler(num => {
              document.getElementById('loader').style.display = 'none';
              if(num) document.getElementById('numero').value = num;
           }).getWhatsAppConfigAtual();
        }
        window.onload = loadNum;
      </script>
    </body></html>`;
  }

  // >>> Template: Gerenciar Cardápio
  else if (templateName === 'GerenciarCardapio') {
    const scriptCRUD = `
      <script>
        let allProducts = []; 
        window.addEventListener('load', loadProducts);

        function showSidebarLoading(isLoading) {
          const el = document.getElementById('sidebar-loader');
          if(el) el.style.display = isLoading ? 'block' : 'none';
        }

        function loadProducts() {
          showSidebarLoading(true);
          google.script.run
            .withSuccessHandler(onProductsLoaded)
            .withFailureHandler(onSidebarError)
             .getCardapioParaAdmin();
        }

        function onProductsLoaded(products) {
          allProducts = products;
          renderList(products);
          showSidebarLoading(false);
        }
        
        function onSidebarError(err) {
          showSidebarLoading(false);
          document.getElementById('product-list').innerHTML = '<p class="text-red-500">Erro: ' + err.message + '</p>';
        }

        function renderList(products) {
          const listEl = document.getElementById('product-list');
          listEl.innerHTML = ''; 
          if (products.length === 0) {
            listEl.innerHTML = '<p class="text-gray-500">Nenhum produto.</p>';
            return;
          }
          products.forEach(p => {
            const color = p.ativo === 'SIM' ? 'text-green-600' : 'text-red-600';
            listEl.innerHTML += \`
               <div class="bg-white p-3 border rounded shadow-sm mb-2 flex justify-between items-center">
                  <div>
                    <div class="font-bold text-gray-800">\${p.nome}</div>
                    <div class="text-xs text-gray-500">\${p.categoria} | R$ \${p.preco.toFixed(2)}</div>
                    <div class="text-xs \${color} font-bold">\${p.ativo}</div>
                  </div>
                  <button onclick="openEditModal('\${p.id}')" class="text-indigo-600 border border-indigo-200 px-2 py-1 rounded hover:bg-indigo-50 font-medium text-sm">Editar</button>
               </div>\`;
          });
        }
        
        document.getElementById('search-box').addEventListener('input', (e) => {
          const term = e.target.value.toLowerCase();
          const filtered = allProducts.filter(p => p.nome.toLowerCase().includes(term));
          renderList(filtered);
        });

        // Modal Logic
        const modal = document.getElementById('edit-modal');
        function closeModal() { modal.style.display = 'none'; }
        
        function openEditModal(id) {
           const p = allProducts.find(x => x.id === id);
           if(!p) return;
           document.getElementById('modal-id').value = p.id;
           document.getElementById('modal-rowIndex').value = p.rowIndex;
           document.getElementById('modal-nome').value = p.nome;
           document.getElementById('modal-preco').value = p.preco;
           document.getElementById('modal-categoria').value = p.categoria;
           document.getElementById('modal-descricao').value = p.descricao || '';
           document.getElementById('modal-foto_url').value = p.foto_url || '';
           document.getElementById('modal-ativo').value = p.ativo;
           document.getElementById('modal-msg').textContent = '';
           modal.style.display = 'block';
        }
        
        // Submit
        document.getElementById('modal-form').addEventListener('submit', (e) => {
           e.preventDefault();
           const btn = document.getElementById('modal-submit-btn');
           btn.disabled = true;
           btn.textContent = 'Salvando...';
           
           const data = {};
           new FormData(e.target).forEach((v, k) => data[k] = v);
           
           google.script.run.withSuccessHandler(res => {
              document.getElementById('modal-msg').textContent = res;
              document.getElementById('modal-msg').className = 'text-green-600 mt-2 text-center font-bold';
              btn.disabled = false;
              btn.textContent = 'Salvar Alterações';
              loadProducts();
              setTimeout(closeModal, 1000);
           }).updateProdutoCardapio(data);
        });
      </script>
    `;
    html = `<html>${head}<body class="p-0 bg-gray-50">
      <div id="sidebar-loader" class="loader" style="display: block; margin: 2rem auto;"></div>
      
      <div class="p-3 bg-white border-b sticky top-0 z-10 shadow-sm">
        <h3 class="text-lg font-bold text-gray-800 mb-2">Gerenciar Cardápio</h3>
        <input type="text" id="search-box" class="w-full border p-2 rounded bg-gray-50 focus:bg-white" placeholder="Buscar...">
      </div>
      
      <div id="product-list" class="p-2 space-y-2"></div>
      
      <div id="edit-modal" class="modal">
         <div class="modal-content">
          <div class="flex justify-between items-center mb-4">
            <h4 class="text-lg font-bold">Editar Produto</h4>
            <span onclick="closeModal()" class="cursor-pointer text-2xl">&times;</span>
          </div>
          <form id="modal-form" class="space-y-3">
             <input type="hidden" id="modal-rowIndex" name="rowIndex">
             <input type="hidden" id="modal-id" name="id">
             
             <div><label class="text-xs font-bold text-gray-500">Nome</label>
             <input id="modal-nome" name="nome" class="w-full border p-2 rounded"></div>
             
             <div class="flex gap-2">
               <div class="flex-1"><label class="text-xs font-bold text-gray-500">Preço</label>
               <input id="modal-preco" name="preco" type="number" step="0.01" class="w-full border p-2 rounded"></div>
               
               <div class="flex-1"><label class="text-xs font-bold text-gray-500">Categoria</label>
               <input id="modal-categoria" name="categoria" class="w-full border p-2 rounded"></div>
             </div>
             
             <div><label class="text-xs font-bold text-gray-500">Descrição</label>
             <textarea id="modal-descricao" name="descricao" rows="2" class="w-full border p-2 rounded"></textarea></div>
             
             <div><label class="text-xs font-bold text-gray-500">Foto URL</label>
             <input id="modal-foto_url" name="foto_url" class="w-full border p-2 rounded"></div>
             
             <div><label class="text-xs font-bold text-gray-500">Status</label>
             <select id="modal-ativo" name="ativo" class="w-full border p-2 rounded bg-white">
                <option value="SIM">Ativo</option>
                <option value="NAO">Inativo</option>
             </select></div>
             
             <button id="modal-submit-btn" class="w-full bg-indigo-600 text-white p-3 rounded font-bold hover:bg-indigo-700 mt-2">Salvar Alterações</button>
          </form>
          <div id="modal-msg"></div>
        </div>
      </div>
      ${scriptCRUD}
    </body></html>`;
  }

  // >>> Template: Gerenciar Insumos
  else if (templateName === 'GerenciarInsumos') {
    html = `<html>${head}<body class="p-0 bg-gray-50">
      <div id="loader" class="loader" style="margin-top: 2rem;"></div>
      <div class="p-3 bg-white border-b sticky top-0 z-10 shadow-sm">
         <h3 class="font-bold text-lg text-gray-800 mb-2">Gerenciar Insumos</h3>
         <input id="search" class="w-full border p-2 rounded bg-gray-50 focus:bg-white" placeholder="Buscar...">
      </div>
      <div id="list" class="p-2 space-y-2"></div>
      
      <div id="modal" class="modal">
         <div class="modal-content">
            <div class="flex justify-between mb-4"><h4 class="font-bold">Editar Insumo</h4><span onclick="closeModal()" class="cursor-pointer text-xl">&times;</span></div>
            <form id="form" class="space-y-3">
               <input type="hidden" name="rowIndex" id="rowIndex">
               <input type="hidden" name="id" id="id">
               
               <div><label class="text-xs font-bold text-gray-500">Nome</label>
               <input name="nome" id="nome" class="w-full border p-2 rounded"></div>
               
               <div><label class="text-xs font-bold text-gray-500">Unidade</label>
               <input name="unidade" id="unidade" class="w-full border p-2 rounded"></div>
               
               <div><label class="text-xs font-bold text-gray-500">Status</label>
               <select name="ativo" id="ativo" class="w-full border p-2 rounded bg-white"><option value="SIM">Ativo</option><option value="NAO">Inativo</option></select></div>
               
               <button class="w-full bg-indigo-600 text-white p-3 rounded font-bold hover:bg-indigo-700 mt-2">Salvar</button>
            </form>
         </div>
      </div>

      <script>
         let allData = [];
         function load() {
            google.script.run.withSuccessHandler(d => {
               allData = d;
               document.getElementById('loader').style.display = 'none';
               render(d);
            }).getInsumosParaAdmin();
         }
         function render(data) {
            const l = document.getElementById('list');
            l.innerHTML = '';
            if(data.length === 0) l.innerHTML = '<p class="text-gray-500 p-2">Nenhum insumo.</p>';
            data.forEach(i => {
               const color = i.ativo === 'SIM' ? 'text-green-600' : 'text-red-600';
               l.innerHTML += \`<div class="bg-white p-3 rounded border shadow-sm flex justify-between items-center">
                  <div>
                     <div class="font-bold text-gray-800">\${i.nome}</div>
                     <div class="text-xs text-gray-500">Estoque: \${i.estoqueAtual} \${i.unidade}</div>
                     <div class="text-xs \${color} font-bold">\${i.ativo}</div>
                  </div>
                  <button onclick="edit('\${i.id}')" class="text-indigo-600 border border-indigo-200 px-2 py-1 rounded hover:bg-indigo-50 font-bold text-sm">Editar</button>
               </div>\`;
            });
         }
         document.getElementById('search').addEventListener('input', (e) => {
            const t = e.target.value.toLowerCase();
            render(allData.filter(x => x.nome.toLowerCase().includes(t)));
         });
         
         const modal = document.getElementById('modal');
         function closeModal() { modal.style.display = 'none'; }
         function edit(id) {
            const i = allData.find(x => x.id === id);
            if(!i) return;
            document.getElementById('rowIndex').value = i.rowIndex;
            document.getElementById('id').value = i.id;
            document.getElementById('nome').value = i.nome;
            document.getElementById('unidade').value = i.unidade;
            document.getElementById('ativo').value = i.ativo;
            modal.style.display = 'block';
         }
         
         document.getElementById('form').addEventListener('submit', (e) => {
            e.preventDefault();
            const btn = e.target.querySelector('button');
            btn.textContent = 'Salvando...';
            btn.disabled = true;
            
            const data = {};
            new FormData(e.target).forEach((v,k)=>data[k]=v);
            
            google.script.run.withSuccessHandler(() => { 
               btn.textContent = 'Salvar'; 
               btn.disabled = false;
               closeModal(); 
               load(); 
            }).updateInsumo(data);
         });
         window.onload = load;
      </script>
    </body></html>`;
  }

  // >>> Template: Gerenciar Taxas (NOVO e CORRIGIDO)
  else if (templateName === 'GerenciarTaxas') {
    html = `<html>${head}<body class="space-y-4">
        <h3 class="text-lg font-semibold text-gray-800">Taxas de Entrega</h3>
        
        <form id="main-form" class="p-3 bg-white border rounded shadow-sm space-y-3">
           <div>
             <label class="block text-sm font-medium text-gray-700">Nome do Bairro</label>
             <input name="bairro" required placeholder="Ex: Centro" class="w-full border p-2 rounded">
           </div>
           <div>
             <label class="block text-sm font-medium text-gray-700">Valor da Taxa (R$)</label>
             <input name="taxa" type="number" step="0.01" required placeholder="Ex: 5.00" class="w-full border p-2 rounded">
           </div>
           <button id="btn-submit" class="w-full bg-green-600 text-white p-2 rounded font-bold hover:bg-green-700">Salvar Taxa</button>
        </form>
        <div id="message"></div>
        <div id="loader" class="loader hidden"></div>

        <div class="mt-4">
           <h4 class="font-bold text-gray-700 mb-2">Bairros Cadastrados</h4>
           <ul id="lista-taxas" class="space-y-2 text-sm">Carregando...</ul>
        </div>

        ${getScriptForm('salvarTaxaEntrega')}
        <script>
           function loadTaxas() {
              google.script.run.withSuccessHandler(renderList).getTaxasParaAdmin();
           }
           function renderList(lista) {
              const ul = document.getElementById('lista-taxas');
              ul.innerHTML = '';
              if(lista.length === 0) ul.innerHTML = '<li class="text-gray-500 italic">Nenhuma taxa cadastrada.</li>';
              
              lista.forEach(item => {
                 const li = document.createElement('li');
                 li.className = 'flex justify-between items-center bg-white p-2 rounded border shadow-sm';
                 li.innerHTML = \`
                    <span><strong>\${item.bairro}</strong>: R$ \${item.taxa.toFixed(2)}</span>
                    <button onclick="excluir('\${item.bairro}')" class="btn-delete">EXCLUIR</button>
                 \`;
                 ul.appendChild(li);
              });
           }
           function excluir(bairro) {
              if(!confirm('Excluir taxa do bairro ' + bairro + '?')) return;
              google.script.run.withSuccessHandler(() => { loadTaxas(); }).excluirTaxaEntrega(bairro);
           }
           window.onload = loadTaxas;
        </script>
      </body></html>`;
  }

  // >>> Template: Ligar Adicionais (CORRIGIDO)
  else if (templateName === 'LigarAdicionais') {
    html = `<html>${head}<body>
      <h3 class="font-bold text-lg mb-4 text-gray-800">Vincular Adicionais</h3>
      <form id="main-form" class="space-y-3">
        <label class="block text-sm font-medium text-gray-700">1. Produto</label>
        <select id="idProduto" name="idProduto" required class="w-full border p-2 rounded" onchange="loadLinks()"></select>
        
        <label class="block text-sm font-medium text-gray-700">2. Selecione os permitidos:</label>
        <div id="lista-checks" class="space-y-2 max-h-60 overflow-y-auto bg-white border p-2 rounded p-2">
           <p class="text-gray-500 text-sm">Carregando...</p>
        </div>
        
        <button id="btn-submit" disabled class="w-full bg-indigo-600 text-white p-2 rounded hover:bg-indigo-700 disabled:opacity-50">Salvar Vínculos</button>
      </form>
      <div id="loader" class="loader hidden mt-4"></div>
      <div id="message"></div>
      
      ${getScriptForm('processarLinkAdicionais')}
      
      <script>
         let todosAdds = [];
         
         function loadBasics() {
            google.script.run.withSuccessHandler(d => {
               const s = document.getElementById('idProduto');
               s.innerHTML = '<option value="" disabled selected>Selecione Produto...</option>';
               d.produtos.forEach(p => s.innerHTML += \`<option value="\${p.id}">\${p.nome}</option>\`);
               todosAdds = d.adicionais;
               document.getElementById('lista-checks').innerHTML = '<p class="text-gray-500 text-sm">Selecione um produto acima.</p>';
            }).getDadosParaLigarAdicionais();
         }
         
         function loadLinks() {
            const idProd = document.getElementById('idProduto').value;
            if(!idProd) return;
            
            const div = document.getElementById('lista-checks');
            div.innerHTML = '<div class="loader"></div>';
            document.getElementById('btn-submit').disabled = true;
            
            google.script.run.withSuccessHandler(linkedIds => {
               div.innerHTML = '';
               if(todosAdds.length === 0) { div.innerHTML = '<p class="text-red-500 text-sm">Sem adicionais cadastrados.</p>'; return; }
               
               todosAdds.forEach(add => {
                  const checked = linkedIds.includes(add.id) ? 'checked' : '';
                  div.innerHTML += \`
                    <label class="checkbox-label">
                      <input type="checkbox" value="\${add.id}" class="chk-addon form-checkbox" \${checked}>
                      <span class="text-sm">\${add.nome} (+R$ \${add.preco.toFixed(2)})</span>
                    </label>\`;
               });
               document.getElementById('btn-submit').disabled = false;
            }).getLinksAtuais(idProd);
         }
         
         window.onload = loadBasics;
      </script>
    </body></html>`;
  }

  else if (templateName === 'VendaBalcao') {
    html = `<html>${head}<body class="bg-gray-100 p-2">
       <div class="bg-white p-4 rounded shadow-lg max-w-md mx-auto">
         <h3 class="text-xl font-bold text-gray-800 mb-4 border-b pb-2">PDV - Balcão</h3>
         
         <div class="mb-4 grid grid-cols-2 gap-2">
            <input id="cli-nome" class="border p-2 rounded text-sm" placeholder="Nome Cliente (Opcional)">
            <select id="cli-pag" class="border p-2 rounded text-sm bg-white"></select>
         </div>

         <div class="bg-gray-50 p-3 rounded border mb-4">
            <label class="text-xs font-bold text-gray-500">ADICIONAR PRODUTO</label>
            <select id="sel-prod" class="w-full border p-2 rounded mb-2 bg-white"></select>
            <div class="flex gap-2">
               <input id="qtd" type="number" value="1" min="1" class="w-20 border p-2 rounded text-center">
               <button onclick="addItem()" class="flex-1 bg-blue-600 text-white font-bold rounded hover:bg-blue-700">+ Adicionar</button>
            </div>
         </div>

         <div id="cart-list" class="space-y-2 max-h-60 overflow-y-auto mb-4 border-t pt-2">
            <p class="text-center text-gray-400 text-sm py-4">Carrinho vazio.</p>
         </div>

         <div class="flex justify-between items-center text-xl font-bold text-gray-800 border-t pt-4 mb-4">
            <span>Total:</span>
            <span id="total">R$ 0.00</span>
         </div>

         <button onclick="finalizar()" id="btn-fin" class="w-full bg-green-600 text-white py-3 rounded font-bold text-lg hover:bg-green-700 shadow">FINALIZAR VENDA</button>
       </div>

       <script>
          let cardapio = [];
          let carrinho = [];

          function init() {
             // Carrega Cardápio
             google.script.run.withSuccessHandler(data => {
                cardapio = data;
                const sel = document.getElementById('sel-prod');
                sel.innerHTML = '<option value="">Selecione...</option>';
                data.forEach(p => {
                   if(p.ativo === 'SIM') {
                       const opt = document.createElement('option');
                       opt.value = p.id;
                       opt.textContent = p.nome + ' - R$ ' + p.preco.toFixed(2);
                       sel.appendChild(opt);
                   }
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

  // >>> Template: Gerenciar Motoqueiros
  else if (templateName === 'GerenciarMotoqueiros') {
    html = `<html>${head}<body>
      <h3 class="font-bold text-lg mb-4 text-gray-800 flex items-center gap-2">
        <span class="material-icons">two_wheeler</span> Gerenciar Motoqueiros
      </h3>
      
      <form id="main-form" class="space-y-3 bg-white p-4 border rounded shadow-sm">
        <div>
          <label class="block text-xs font-bold text-gray-500 uppercase">Nome do Entregador</label>
          <input name="nome" placeholder="Ex: João Silva" required class="w-full border p-2 rounded focus:ring-2 focus:ring-blue-500 outline-none">
        </div>
        <div class="grid grid-cols-2 gap-2">
           <div>
             <label class="block text-xs font-bold text-gray-500 uppercase">WhatsApp</label>
             <input name="telefone" placeholder="21 99999-9999" class="w-full border p-2 rounded">
           </div>
           <div>
             <label class="block text-xs font-bold text-gray-500 uppercase">Placa/Veículo</label>
             <input name="placa" placeholder="ABC-1234" class="w-full border p-2 rounded">
           </div>
        </div>
        <button id="btn-submit" class="w-full bg-blue-600 text-white p-2 rounded font-bold hover:bg-blue-700 transition">
          CADASTRAR MOTOQUEIRO
        </button>
      </form>
      
      <div id="msg" class="mt-2"></div>
      
      <div class="mt-6">
        <h4 class="font-bold text-gray-700 mb-2 border-b pb-1">Entregadores Ativos</h4>
        <div id="list" class="space-y-2">
           <div class="text-center text-gray-400 py-4"><div class="loader"></div> Carregando...</div>
        </div>
      </div>

      ${getScriptForm('salvarMotoqueiro')}
      
      <script>
        function loadMotos(){
          google.script.run.withSuccessHandler(l => {
             const list = document.getElementById('list');
             if(l.length === 0) {
               list.innerHTML = '<p class="text-gray-500 text-sm text-center bg-gray-50 p-3 rounded">Nenhum motoqueiro cadastrado.</p>';
               return;
             }
             list.innerHTML = l.map(m => \`
               <div class="bg-white p-3 border rounded shadow-sm flex justify-between items-center group hover:border-blue-300 transition">
                 <div>
                   <div class="font-bold text-gray-800 flex items-center gap-2">
                      <span class="material-icons text-sm text-gray-400">person</span> \${m.nome}
                   </div>
                   <div class="text-xs text-gray-500 ml-5">
                     \${m.telefone ? '📞 '+m.telefone : ''} \${m.placa ? ' | 🛵 '+m.placa : ''}
                   </div>
                 </div>
                 <button onclick="del('\${m.id}')" class="text-red-400 hover:text-red-600 hover:bg-red-50 p-2 rounded transition" title="Excluir">
                    <span class="material-icons text-lg">delete</span>
                 </button>
               </div>
             \`).join('');
          }).getMotoqueirosAdmin();
        } 
        
        function del(id){
          if(confirm('Tem certeza que deseja remover este entregador?')) {
             google.script.run.withSuccessHandler(() => {
                loadMotos();
             }).excluirMotoqueiro(id);
          }
        }
        
        // Hook para recarregar após salvar (usando a função m() do script padrão)
        const originalM = m; // Backup da função de mensagem
        m = function(t, e) {
           originalM(t, e); // Chama original
           if(!e) loadMotos(); // Se não for erro, recarrega lista
        };

        window.onload = loadMotos;
      </script>
    </body></html>`;
  }

  // Fallback (404)
  else {
    html = `<html>${head}<body>
      <h3 class="text-red-600">Erro 404: Template "${templateName}" não encontrado.</h3>
      <p>Verifique o nome do arquivo no Código.gs.</p>
    </body></html>`;
  }

  return HtmlService.createHtmlOutput(html).setSandboxMode(HtmlService.SandboxMode.IFRAME);
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