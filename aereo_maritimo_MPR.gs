/**
 * @OnlyCurrentDoc false
 */

const CACHE_MPR = {
  planilhaPrimaria: null,
  mapeamentoIndices: null,
  timestampCache: null,
  pastasCache: {},
  arquivosProcessados: {},
  mapaArea: {
    "(0019-00) - MPR": "MPRQB19",
  },
  ultimaLimpeza: null,
  // Configuração de cache
  TEMPO_CACHE_PLANILHA: 30 * 60 * 1000, // 30 minutos
  TEMPO_CACHE_ARQUIVOS: 24 * 60 * 60 * 1000 // 24 horas
};

// Configurações para cada tipo de aba
const CONFIG_ABAS = {
  "VCP": {
    colunaGatilho: 1,
    colunas: {
      n_pedido: "B",
      ref_nh: "C", 
      numero_di: "D",
      fornecedor: "E",
      fatura: "F",
      master: "G",
      house: "H",
      volumes: "I",
      pb: "J",
      m3: "K",
      cif: "L",
    },
    processarDI: false,
    processarCI: true,
    tipo: "AEREO"
  },
  "GRU": {
    colunaGatilho: 1,
    colunas: {
      n_pedido: "B",
      ref_nh: "C", 
      numero_di: "D",
      fornecedor: "E",
      fatura: "F",
      master: "G",
      house: "H",
      volumes: "I",
      pb: "J",
      m3: "K",
      cif: "L",
    },
    processarDI: false,
    processarCI: true,
    tipo: "AEREO"
  },
  "CWB": {
    colunaGatilho: 1,
    colunas: {
      n_pedido: "B",
      ref_nh: "C", 
      numero_di: "D",
      fornecedor: "E",
      fatura: "F",
      master: "G",
      house: "H",
      volumes: "I",
      pb: "J",
      m3: "K",
      cif: "L",
    },
    processarDI: false,
    processarCI: true,
    tipo: "AEREO"
  },
  "PNG MARÍTIMO": {
    colunaGatilho: 1,
    colunas: {
      n_pedido: "B",
      ref_nh: "C",
      area: "D",
      numero_di: "E", 
      fornecedor: "F",
      fatura: "G",
      cntr: "H",
      bl: "I",
      volumes: "J",
      pb: "K",
      m3: "L",
      cif: "M",
      valores_deim: "N",
      afrmm: "O",
      thd: "P",
      li: "Q"
    },
    processarDI: false,
    processarCI: true,
    tipo: "MARITIMO"
  },
  "RODOVIÁRIO": {
    colunaGatilho: 1,
    colunas: {
      transportadora: "B",
      crt: "C",
      fatura: "D",
      n_pedido: "E", 
      canal: "F",
      di: "G",
      cavalo: "H",
      carreta: "I",
    },
    processarDI: false,
    processarCI: false,
    tipo: "RODOVIARIO"
  }
};

function obterConfiguracoes() {
  const properties = PropertiesService.getScriptProperties();
  return {
    planilhaId: properties.getProperty('PLANILHA_ID') || "1C9ZkOqXBh5IdjXP8K8CuEsmmQgy5C1yZxXHKJFnHj24",
    pastaId: properties.getProperty('PASTA_ID') || "1NjYLk2dt8PB3RpaJ0C9O8K_Gs_8OIHoZ"
  };
}

// NOVA FUNÇÃO: Pré-carregar planilha primária com cache inteligente
function preCarregarPlanilhaPrimaria(forcarRecarregamento = false) {
  const agora = new Date().getTime();
  
  // Verificar se precisa recarregar
  const precisaRecarregar = forcarRecarregamento || 
    !CACHE_MPR.planilhaPrimaria || 
    !CACHE_MPR.timestampCache || 
    (agora - CACHE_MPR.timestampCache) > CACHE_MPR.TEMPO_CACHE_PLANILHA;
  
  if (!precisaRecarregar) {
    Logger.log("✓ Usando planilha primária do cache");
    return {
      dados: CACHE_MPR.planilhaPrimaria,
      mapeamento: CACHE_MPR.mapeamentoIndices
    };
  }
  
  Logger.log("📊 Carregando planilha primária...");
  
  try {
    const config = obterConfiguracoes();
    const ss = SpreadsheetApp.openById(config.planilhaId);
    const sheet = ss.getSheetByName("ANDAM.");
    
    if (!sheet) {
      throw new Error("Planilha 'ANDAM.' não encontrada");
    }
    
    Logger.log(`Planilha encontrada: ${sheet.getName()}`);
    
    // Carregar todos os dados de uma vez
    const dados = sheet.getDataRange().getValues();
    
    // Criar mapeamento de índices do cabeçalho
    const mapeamento = obterIndicesDaColunas(sheet);
    
    // Atualizar cache
    CACHE_MPR.planilhaPrimaria = dados;
    CACHE_MPR.mapeamentoIndices = mapeamento;
    CACHE_MPR.timestampCache = agora;
    
    Logger.log(`✓ Planilha primária carregada com ${dados.length - 1} registros`);
    Logger.log(`✓ Cache atualizado com timestamp: ${new Date(agora).toLocaleString()}`);
    
    return {
      dados: dados,
      mapeamento: mapeamento
    };
    
  } catch (error) {
    Logger.log(`❌ Erro ao carregar planilha primária: ${error}`);
    throw error;
  }
}

// FUNÇÃO ATUALIZADA: Buscar múltiplos IDs de uma vez
function buscarDadosPorIds(ids) {
  const { dados } = preCarregarPlanilhaPrimaria();
  const resultados = {};
  
  Logger.log(`🔍 Buscando dados para ${ids.length} IDs: ${ids.join(', ')}`);
  
  // Buscar todos os IDs de uma vez
  for (let i = 1; i < dados.length; i++) { // Começar de 1 para pular cabeçalho
    const linha = dados[i];
    const idLinha = linha[0]?.toString().trim();
    
    if (idLinha && ids.includes(idLinha)) {
      resultados[idLinha] = linha;
      Logger.log(`✓ Encontrado: ${idLinha}`);
    }
  }
  
  // Verificar quais IDs não foram encontrados
  const naoEncontrados = ids.filter(id => !resultados[id]);
  if (naoEncontrados.length > 0) {
    Logger.log(`⚠️ IDs não encontrados: ${naoEncontrados.join(', ')}`);
  }
  
  return resultados;
}

function obterIndicesDaColunas(sheet) {
  const cabecalho = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const mapa = {};
  
  cabecalho.forEach((valor, indice) => {
    const nomeColuna = valor.toString().trim().toUpperCase();
    mapa[nomeColuna] = indice;
  });
  
  return mapa;
}

// FUNÇÃO OTIMIZADA: Limpar cache com controle mais granular
function limparCacheSeNecessario() {
  const agora = new Date().getTime();
  
  if (!CACHE_MPR.ultimaLimpeza || (agora - CACHE_MPR.ultimaLimpeza) > CACHE_MPR.TEMPO_CACHE_ARQUIVOS) {
    Logger.log("🧹 Iniciando limpeza de cache...");
    
    let arquivosLimpos = 0;
    // Limpar cache de arquivos antigos
    for (const key in CACHE_MPR.arquivosProcessados) {
      if (CACHE_MPR.arquivosProcessados[key].timestamp && 
          (agora - CACHE_MPR.arquivosProcessados[key].timestamp) > CACHE_MPR.TEMPO_CACHE_ARQUIVOS) {
        delete CACHE_MPR.arquivosProcessados[key];
        arquivosLimpos++;
      }
    }
    
    // Limpar cache de pastas também
    CACHE_MPR.pastasCache = {};
    CACHE_MPR.ultimaLimpeza = agora;
    
    Logger.log(`✓ Cache limpo: ${arquivosLimpos} arquivos removidos`);
  }
  
  // Verificar se planilha primária precisa ser atualizada
  if (CACHE_MPR.timestampCache && 
      (agora - CACHE_MPR.timestampCache) > CACHE_MPR.TEMPO_CACHE_PLANILHA) {
    Logger.log("⏰ Cache da planilha primária expirado, será recarregado na próxima chamada");
  }
}

// FUNÇÃO MELHORADA: Validação de dados com logging detalhado
function validarDados(linhaDados, indices) {
  if (!linhaDados || !Array.isArray(linhaDados)) {
    Logger.log("❌ Dados da linha são inválidos ou não é array");
    return false;
  }
  
  const indicesInvalidos = indices.filter(indice => 
    indice >= linhaDados.length || linhaDados[indice] === undefined
  );
  
  if (indicesInvalidos.length > 0) {
    Logger.log(`❌ Índices inválidos encontrados: ${indicesInvalidos.join(', ')} (linha tem ${linhaDados.length} colunas)`);
    return false;
  }
  
  return true;
}

// FUNÇÃO SIMPLIFICADA: Agora usa o pré-carregamento
function getDadosPrimaria() {
  const { dados, mapeamento } = preCarregarPlanilhaPrimaria();
  return dados;
}

// FUNÇÃO PRINCIPAL OTIMIZADA
function gatilho_MPR(e) {
  const inicioExecucao = new Date().getTime();
  Logger.log("🚀 Iniciando gatilho_MPR...");
  
  try {
    // Limpar cache se necessário
    limparCacheSeNecessario();
    
    const sheet = e.source.getActiveSheet();
    const range = e.range;
    const nomeAba = sheet.getName();
    
    Logger.log(`📋 Aba detectada: ${nomeAba}`);
    
    // Verificar se é uma aba configurada e se a coluna editada é a correta
    const config = CONFIG_ABAS[nomeAba];
    if (!config || range.getColumn() !== config.colunaGatilho) {
      Logger.log("⏭️ Aba não configurada ou coluna incorreta. Saindo.");
      return;
    }

    Logger.log(`📍 Linha editada: ${range.getRow()}, Coluna: ${range.getColumn()}`);
    
    // PRÉ-CARREGAR a planilha primária uma única vez
    preCarregarPlanilhaPrimaria();
    
    const configuracoes = obterConfiguracoes();
    const pastaRaiz = DriveApp.getFolderById(configuracoes.pastaId);
    
    // Coletar todos os IDs primeiro
    const idsParaBuscar = [];
    range.getValues().forEach((row) => {
      const idInserido = row[0].toString().trim();
      if (idInserido) {
        idsParaBuscar.push(idInserido);
      }
    });
    
    if (idsParaBuscar.length === 0) {
      Logger.log("⚠️ Nenhum ID válido encontrado");
      return;
    }
    
    // Buscar todos os dados de uma vez
    const dadosEncontrados = buscarDadosPorIds(idsParaBuscar);
    
    // Coletar todas as atualizações primeiro
    const updates = [];
    
    range.getValues().forEach((row, i) => {
      const idInserido = row[0].toString().trim();
      
      if (!idInserido) return;
      
      const linhaDados = dadosEncontrados[idInserido];
      if (!linhaDados) {
        Logger.log(`❌ Dados não encontrados para ID: ${idInserido}`);
        return;
      }

      // Validar dados antes de processar
      const indicesNecessarios = [1, 5, 12, 13, 253, 456, 457, 458, 459, 461, 471];
      if (!validarDados(linhaDados, indicesNecessarios)) {
        Logger.log(`❌ Dados inválidos para ID: ${idInserido}`);
        return;
      }

      const linha = range.getRow() + i;
      const fatura = linhaDados[457];
      const pedidoConvertido = converterFaturaParaPedido(idInserido, fatura);

      // Processar dados conforme o tipo de aba
      if (config.tipo === "AEREO") {
        updates.push({
          linha,
          valores: processarDadosAereo(linhaDados, pedidoConvertido, nomeAba),
          idInserido,
          config,
          linhaDados,
          tipoAba: nomeAba
        });
      } else if (config.tipo === "MARITIMO") {
        updates.push({
          linha,
          valores: processarDadosMaritimo(linhaDados, pedidoConvertido),
          idInserido,
          config,
          linhaDados,
          tipoAba: nomeAba
        });
      }
    });
    
    Logger.log(`📝 Processando ${updates.length} atualizações...`);
    
    // Processar todas as atualizações
    updates.forEach((update, index) => {
      Logger.log(`📄 Processando atualização ${index + 1}/${updates.length} - ID: ${update.idInserido}`);
      
      preencherDadosBasicos(sheet, update, update.tipoAba);
      
      const subpasta = findSubpasta(pastaRaiz, update.idInserido);
      if (subpasta) {
        Logger.log(`📁 Subpasta encontrada: ${subpasta.getName()}`);
        
        if (update.config.tipo === "MARITIMO") {
          processarDocumentosMaritimo(subpasta, update.idInserido, sheet, update.linha, update.config);
        } else {
          processarDocumentosAereo(subpasta, update.idInserido, sheet, update.linha, update.config, update.tipoAba);
        }
      } else {
        Logger.log(`❌ Subpasta não encontrada para ID: ${update.idInserido}`);
      }
    });
    
    const tempoExecucao = new Date().getTime() - inicioExecucao;
    Logger.log(`✅ Gatilho_MPR concluído em ${tempoExecucao}ms`);
    
  } catch (err) {
    const tempoExecucao = new Date().getTime() - inicioExecucao;
    Logger.log(`❌ Erro na função gatilho_MPR após ${tempoExecucao}ms: ${err}`);
    Logger.log(`📚 Stack trace: ${err.stack}`);
  }
}

// NOVA FUNÇÃO: Pré-carregar cache manualmente (para testes)
function preCarregarCacheManual() {
  try {
    Logger.log("🔄 Iniciando pré-carregamento manual do cache...");
    const resultado = preCarregarPlanilhaPrimaria(true); // Forçar recarregamento
    Logger.log(`✅ Cache pré-carregado com sucesso: ${resultado.dados.length - 1} registros`);
    return true;
  } catch (error) {
    Logger.log(`❌ Erro no pré-carregamento: ${error}`);
    return false;
  }
}

// NOVA FUNÇÃO: Verificar status do cache
function verificarStatusCache() {
  const agora = new Date().getTime();
  const status = {
    planilhaPrimaria: {
      carregada: !!CACHE_MPR.planilhaPrimaria,
      registros: CACHE_MPR.planilhaPrimaria ? CACHE_MPR.planilhaPrimaria.length - 1 : 0,
      ultimaAtualizacao: CACHE_MPR.timestampCache ? new Date(CACHE_MPR.timestampCache).toLocaleString() : 'Nunca',
      expirado: CACHE_MPR.timestampCache ? (agora - CACHE_MPR.timestampCache) > CACHE_MPR.TEMPO_CACHE_PLANILHA : true
    },
    arquivos: {
      quantidade: Object.keys(CACHE_MPR.arquivosProcessados).length,
      ultimaLimpeza: CACHE_MPR.ultimaLimpeza ? new Date(CACHE_MPR.ultimaLimpeza).toLocaleString() : 'Nunca'
    },
    pastas: {
      quantidade: Object.keys(CACHE_MPR.pastasCache).length
    }
  };
  
  Logger.log("📊 STATUS DO CACHE:");
  Logger.log(`   Planilha Primária: ${status.planilhaPrimaria.carregada ? 'CARREGADA' : 'NÃO CARREGADA'}`);
  Logger.log(`   Registros: ${status.planilhaPrimaria.registros}`);
  Logger.log(`   Última Atualização: ${status.planilhaPrimaria.ultimaAtualizacao}`);
  Logger.log(`   Status: ${status.planilhaPrimaria.expirado ? 'EXPIRADO' : 'VÁLIDO'}`);
  Logger.log(`   Arquivos em Cache: ${status.arquivos.quantidade}`);
  Logger.log(`   Pastas em Cache: ${status.pastas.quantidade}`);
  
  return status;
}


function extrairTextoDoPDF(arquivo) {
  const blob = arquivo.getBlob();
  const resource = {
    title: blob.getName(),
    mimeType: blob.getContentType()
  };
  
  try {
    const file = Drive.Files.insert(resource, blob, {ocr: true, ocrLanguage: "pt"});
    const doc = DocumentApp.openById(file.id);
    const text = doc.getBody().getText();
    Drive.Files.remove(file.id);
    return text;
  } catch (e) {
    Logger.log(`Erro ao extrair texto do PDF: ${e}`);
    return "";
  }
}

function processarArquivo(arquivo, idInserido) {
  try {
    const nome = arquivo.getName().toUpperCase();
    
    // Verificar se é um PDF
    if (!arquivo.getName().toLowerCase().endsWith('.pdf')) {
      Logger.log(`Arquivo não é PDF: ${arquivo.getName()}`);
      return null;
    }
    
    const texto = extrairTextoDoPDF(arquivo);
    
    if (!texto || texto.length < 10) {
      Logger.log(`Texto extraído é muito curto ou vazio para: ${arquivo.getName()}`);
      return null;
    }
    
    let resultado = null;
    
    if (nome.includes('_CI_')) {
      Logger.log('Identificado como CI - extraindo dados completos');
      resultado = extrairCIFCI(texto);
      if (resultado) {
        resultado.tipoDocumento = 'CI';
        Logger.log(`Dados CI extraídos: ${JSON.stringify(resultado)}`);
      }
    }
    else if (nome.includes('_DI_')) {
      Logger.log('Identificado como DI - extraindo campos específicos');
      resultado = extrairDadosDI(texto);
      if (resultado) {
        resultado.tipoDocumento = 'DI';
        Logger.log(`Dados DI extraídos: ${JSON.stringify(resultado)}`);
      }
    }
    
    if (resultado) {
      resultado.timestamp = new Date().getTime();
      resultado.nomeArquivo = arquivo.getName();
    }
    
    return resultado;
    
  } catch (e) {
    Logger.log(`Erro no processamento: ${e.toString()}`);
    return null;
  }
}

// Função para processar dados aéreos
function processarDadosAereo(linhaDados, pedidoConvertido, nomeAba) {
  return [
    pedidoConvertido,                    // N_PEDIDO (B)
    linhaDados[1] || "",                 // REF_NH (C)
    linhaDados[253].toString().replace(/[-\/]/g, ""),               // NUMERO_DI (D)
    linhaDados[471] || "",                 // FORNECEDOR (E)
    linhaDados[457] || "",               // FATURA (F)
    linhaDados[12] || "",                 // MASTER (G)
    linhaDados[13] || "",                  // HOUSE (H)
    "",                                  // VOLUMES (I)
    "",                                  // PB (J) -
    "",                                  // M3 (K) -
    "",                                  // CIF (L) -
    linhaDados[461] || ""                  // VALORES DEIM (M)
  ];
}

// Função para processar dados marítimos
function buscarCotacaoDolar() {
  try {
    // Data atual no formato MM-dd-yyyy (formato exigido pela API do BC)
    const hoje = new Date();
    const dataFormatada = `${String(hoje.getMonth() + 1).padStart(2, '0')}-${String(hoje.getDate()).padStart(2, '0')}-${hoje.getFullYear()}`;
    
    // URL da API do Banco Central para dólar comercial
    const url = `https://olinda.bcb.gov.br/olinda/servico/PTAX/versao/v1/odata/CotacaoDolarDia(dataCotacao=@dataCotacao)?@dataCotacao='${dataFormatada}'&$top=1&$format=json`;
    
    Logger.log(`Buscando cotação do dólar para: ${dataFormatada}`);
    Logger.log(`URL: ${url}`);
    
    const response = UrlFetchApp.fetch(url);
    const data = JSON.parse(response.getContentText());
    
    if (data.value && data.value.length > 0) {
      const cotacao = data.value[0].cotacaoVenda;
      Logger.log(`✓ Cotação encontrada: R$ ${cotacao}`);
      return cotacao;
    } else {
      Logger.log('⚠️ Cotação não encontrada para hoje, tentando dia anterior...');
      return buscarCotacaoDolarDiaAnterior();
    }
    
  } catch (error) {
    Logger.log(`❌ Erro ao buscar cotação: ${error.message}`);
    return buscarCotacaoDolarDiaAnterior();
  }
}

// Função auxiliar para buscar cotação do dia anterior (caso hoje não tenha)
function buscarCotacaoDolarDiaAnterior() {
  try {
    // Data de ontem
    const ontem = new Date();
    ontem.setDate(ontem.getDate() - 1);
    const dataFormatada = `${String(ontem.getMonth() + 1).padStart(2, '0')}-${String(ontem.getDate()).padStart(2, '0')}-${ontem.getFullYear()}`;
    
    const url = `https://olinda.bcb.gov.br/olinda/servico/PTAX/versao/v1/odata/CotacaoDolarDia(dataCotacao=@dataCotacao)?@dataCotacao='${dataFormatada}'&$top=1&$format=json`;
    
    Logger.log(`Tentando cotação de ontem: ${dataFormatada}`);
    
    const response = UrlFetchApp.fetch(url);
    const data = JSON.parse(response.getContentText());
    
    if (data.value && data.value.length > 0) {
      const cotacao = data.value[0].cotacaoVenda;
      Logger.log(`✓ Cotação de ontem encontrada: R$ ${cotacao}`);
      return cotacao;
    } else {
      Logger.log('⚠️ Usando cotação padrão: R$ 5.50');
      return 5.50; // Valor padrão caso não encontre
    }
    
  } catch (error) {
    Logger.log(`❌ Erro ao buscar cotação de ontem: ${error.message}`);
    Logger.log('⚠️ Usando cotação padrão: R$ 5.50');
    return 5.50; // Valor padrão
  }
}

// Função para armazenar cotação no cache para não buscar múltiplas vezes
function obterCotacaoDolarComCache() {
  // Verificar se já temos cotação no cache de hoje
  const hoje = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const chaveCache = `cotacao_usd_${hoje}`;
  
  let cotacao = CacheService.getScriptCache().get(chaveCache);
  
  if (cotacao) {
    Logger.log(`✓ Cotação do cache: R$ ${cotacao}`);
    return parseFloat(cotacao);
  }
  
  // Se não tem no cache, buscar na API
  cotacao = buscarCotacaoDolar();
  
  if (cotacao) {
    // Armazenar no cache por 24 horas
    CacheService.getScriptCache().put(chaveCache, cotacao.toString(), 86400);
    return cotacao;
  }
  
  return 5.50; // Valor padrão
}

// Função processarDadosMaritimo atualizada
function processarDadosMaritimo(linhaDados, pedidoConvertido) {
  // Função auxiliar para converter USD para BRL se necessário
  function converterParaBRL(valor) {
    if (!valor) return 0;
    
    const valorStr = valor.toString().trim();
    
    // Se contém "USD", extrair número e converter
    if (valorStr.includes('USD')) {
      const numeroUSD = parseFloat(valorStr.replace(/[^\d.,]/g, '').replace(',', '.'));
      if (!isNaN(numeroUSD)) {
        const cotacaoUSD = obterCotacaoDolarComCache();
        return numeroUSD * cotacaoUSD;
      }
      return 0;
    }
    
    // Se é apenas número, retornar como está (já em BRL)
    const numeroBRL = parseFloat(valorStr.replace(/[^\d.,]/g, '').replace(',', '.'));
    return isNaN(numeroBRL) ? 0 : numeroBRL;
  }
  
  // Calcular TOTAL VALORES DEIM
  const thdConvertido = converterParaBRL(linhaDados[465] || ""); // THD
  const afrmm = parseFloat((linhaDados[464] || "").toString().replace(/[^\d.,]/g, '').replace(',', '.')) || 0; // AFRMM
  const li = parseFloat((linhaDados[461] || "").toString().replace(/[^\d.,]/g, '').replace(',', '.')) || 0; // LI
  
  const totalValoresDEIM = thdConvertido + afrmm + li;
  
  Logger.log(`Cálculo TOTAL VALORES DEIM:`);
  Logger.log(`THD original: ${linhaDados[465]} | THD convertido: ${thdConvertido}`);
  Logger.log(`AFRMM: ${afrmm} | LI: ${li}`);
  Logger.log(`Total: ${totalValoresDEIM}`);
  
  return [
    pedidoConvertido,                    // N_PEDIDO (B)
    linhaDados[1] || "",                 // REF_NH (C)
    CACHE_MPR.mapaArea[linhaDados[5]] || linhaDados[5],  // AREA (D)
    linhaDados[253].toString().replace(/[-\/]/g, ""), // NUMERO_DI (E)
    linhaDados[471] || "",               // FORNECEDOR (F)
    linhaDados[457] || "",               // FATURA (G)
    linhaDados[13] || "",                // CNTR (H)
    linhaDados[12] || "",                // BL (I)
    "",                                  // VOLUMES (J) - será preenchido pelo processamento
    "",                                  // PB (K) - será preenchido pelo processamento
    "",                                  // M3 (L) - será preenchido pelo processamento
    "",                                  // CIF (M) - será preenchido pelo processamento
    totalValoresDEIM.toFixed(2),         // TOTAL VALORES DEIM (N)
    linhaDados[464] || "",               // AFRMM (O)
    linhaDados[465] || "",               // THD (P)
    linhaDados[461] || ""                // LI (Q)
  ];
}

// Função para converter a fatura em número de pedido baseado na origem da referência
function converterFaturaParaPedido(idInserido, faturaOriginal) {
  if (!faturaOriginal) return "";

  let origem = idInserido.substring(3, 5).toUpperCase();

  return faturaOriginal
    .toString()
    .split("/")
    .map(faturaBruta => {
      const faturaLimpa = faturaBruta.trim().split(".")[0].replace(/[^A-Za-z0-9]/g, "");

      switch (origem) {
        case "FR":
          return "F25P" + faturaLimpa.slice(-6);
        case "RO":
          return "R25P" + faturaLimpa.slice(-6);
        case "MX":
          return "M25P" + faturaLimpa.slice(-6);
        case "US":
          return "F25P" + faturaLimpa.slice(-6);
        case "IN":
          return "I25P" + faturaLimpa.slice(-6);
        case "CO":
          return "C25P" + faturaLimpa.slice(-6);
        case "PY":
          const faturaSemHifens = faturaLimpa.replace(/-/g, "");
          return "P25P" + faturaSemHifens.slice(-6);
        default:
          return faturaOriginal;
      }
    })
    .join(" / ");
}

function preencherDadosBasicos(sheet, update, nomeAba) {
  if (nomeAba === "VCP" || nomeAba === "GRU" || nomeAba === "CWB") {
    // Preencher colunas B-H para aéreo
    const range = sheet.getRange(`B${update.linha}:M${update.linha}`);
    range.setValues([update.valores]);
    Logger.log(`✓ Dados básicos aéreo preenchidos: ${JSON.stringify(update.valores.slice(0, 7))}`);
  }

  if (nomeAba === "PNG MARÍTIMO") {
    // Preencher colunas B-Q para marítimo
    const range = sheet.getRange(`B${update.linha}:Q${update.linha}`);
    range.setValues([update.valores]);
    Logger.log(`✓ Dados básicos marítimo preenchidos: ${JSON.stringify(update.valores)}`);
  }
}

function processarDocumentosAereo(subpasta, idInserido, sheet, linha, config, tipoAba) {
  const arquivos = subpasta.getFiles();
  const resultados = [];
  const idUpper = idInserido.toUpperCase();
  
  let arquivosEncontrados = [];
  
  while (arquivos.hasNext()) {
    const arquivo = arquivos.next();
    const nome = arquivo.getName().toUpperCase();
    arquivosEncontrados.push(nome);
    
    let deveProcessarArquivo = false;
    
    if (config.processarDI && nome.includes(`_DI_${idUpper}.PDF`)) {
      deveProcessarArquivo = true;
      Logger.log(`Arquivo DI encontrado: ${nome}`);
    }
    if (config.processarCI && nome.includes(`_CI_${idUpper}.PDF`)) {
      deveProcessarArquivo = true;
      Logger.log(`Arquivo CI encontrado: ${nome}`);
    }
    
    if (deveProcessarArquivo) {
      const cacheFileKey = `${idUpper}_${nome}`;
      if (!CACHE_MPR.arquivosProcessados[cacheFileKey]) {
        Logger.log(`Processando arquivo: ${nome}`);
        const resultado = processarArquivo(arquivo, idInserido);
        if (resultado) {
          resultado.tipoDocumento = nome.includes('_DI_') ? 'DI' : 'CI';
          resultados.push(resultado);
          CACHE_MPR.arquivosProcessados[cacheFileKey] = resultado;
          Logger.log(`✓ Arquivo processado com sucesso: ${nome}`);
        } else {
          Logger.log(`✗ Falha no processamento do arquivo: ${nome}`);
        }
      } else {
        resultados.push(CACHE_MPR.arquivosProcessados[cacheFileKey]);
        Logger.log(`✓ Usando dados do cache para: ${nome}`);
      }
    }
  }
  
  if (resultados.length === 0) {
    return;
  }
  
  Logger.log(`Processando ${resultados.length} resultados para linha ${linha}`);
  preencherValoresAereo(sheet, linha, resultados, config);
}

function processarDocumentosMaritimo(subpasta, idInserido, sheet, linha, config) {
  Logger.log(`=== PROCESSANDO DOCUMENTOS MARÍTIMO - ID: ${idInserido}, LINHA: ${linha} ===`);
  
  const arquivos = subpasta.getFiles();
  const resultados = [];
  const idUpper = idInserido.toUpperCase();
  
  let arquivosEncontrados = [];
  
  while (arquivos.hasNext()) {
    const arquivo = arquivos.next();
    const nome = arquivo.getName().toUpperCase();
    arquivosEncontrados.push(nome);
    
    let deveProcessarArquivo = false;
    
    if (config.processarDI && nome.includes(`_DI_${idUpper}.PDF`)) {
      deveProcessarArquivo = true;
      Logger.log(`Arquivo DI encontrado: ${nome}`);
    }
    if (config.processarCI && nome.includes(`_CI_${idUpper}.PDF`)) {
      deveProcessarArquivo = true;
      Logger.log(`Arquivo CI encontrado: ${nome}`);
    }
    
    if (deveProcessarArquivo) {
      const cacheFileKey = `${idUpper}_${nome}`;
      if (!CACHE_MPR.arquivosProcessados[cacheFileKey]) {
        Logger.log(`Processando arquivo: ${nome}`);
        const resultado = processarArquivo(arquivo, idInserido);
        if (resultado) {
          resultado.tipoDocumento = nome.includes('_DI_') ? 'DI' : 'CI';
          resultados.push(resultado);
          CACHE_MPR.arquivosProcessados[cacheFileKey] = resultado;
          Logger.log(`✓ Arquivo processado com sucesso: ${nome}`);
        } else {
          Logger.log(`✗ Falha no processamento do arquivo: ${nome}`);
        }
      } else {
        resultados.push(CACHE_MPR.arquivosProcessados[cacheFileKey]);
        Logger.log(`✓ Usando dados do cache para: ${nome}`);
      }
    }
  }
  
  if (resultados.length === 0) {
    return;
  }
  
  Logger.log(`Processando ${resultados.length} resultados para linha ${linha}`);
  preencherValoresMaritimo(sheet, linha, resultados, config);
}

function preencherValoresAereo(sheet, linha, resultados, config) {
 
   // Separar dados do DI e CI
  let dadosDI = null;
  let dadosCI = null;
  
  resultados.forEach(r => {
  if (r.tipoDocumento === 'DI') {
    dadosDI = r;
 
  } else if (r.tipoDocumento === 'CI') {
    dadosCI = r;
    Logger.log(`Dados CI encontrados: ${JSON.stringify(dadosCI)}`);
    }
  });
  
   // Preencher VOLUMES (I) - priorizar DI, depois CI
  let volumes = "";
  if (dadosDI?.QTDE_VOLUMES) {
    volumes = dadosDI.QTDE_VOLUMES;

  } else if (dadosCI?.QTDE_VOLUMES) {
    volumes = dadosCI.QTDE_VOLUMES;
    Logger.log(`VOLUMES do CI: ${volumes}`);
  }
  
 // Preencher PB (J) - priorizar DI, depois CI
  let pb = "";
  if (dadosDI?.PB) {
    pb = formatarParaExibicao(dadosDI.PB, '', 'PB');
 
  } else if (dadosCI?.PB) {
    pb = formatarParaExibicao(dadosCI.PB, '', 'PB');
    Logger.log(`PB do CI: ${pb}`);
  }
 
 // Preencher CIF (L) - priorizar CI, depois DI
  let cif = "";
    if (dadosCI?.CIF) {
    cif = formatarParaExibicao(dadosCI.CIF, 'R$ ', 'CIF');
  Logger.log(`CIF do CI: ${cif}`);
  } else if (dadosDI?.CIF) {
    cif = formatarParaExibicao(dadosDI.CIF, 'R$ ', 'CIF');
 
  }
 
  // Preencher as células
  if (volumes) {
    sheet.getRange(`I${linha}`).setValue(volumes);
    Logger.log(`✓ VOLUMES preenchido na célula I${linha}: ${volumes}`);
  }
  if (pb) {
    sheet.getRange(`J${linha}`).setValue(pb);
    Logger.log(`✓ PB preenchido na célula J${linha}: ${pb}`);
  }
  if (cif) {
    sheet.getRange(`L${linha}`).setValue(cif);
    Logger.log(`✓ CIF preenchido na célula L${linha}: ${cif}`);
  }
 
}

function preencherValoresMaritimo(sheet, linha, resultados, config) {
  // Separar dados do DI e CI
  let dadosDI = null;
  let dadosCI = null;
  
  resultados.forEach(r => {
    if (r.tipoDocumento === 'DI') {
      dadosDI = r;
    } else if (r.tipoDocumento === 'CI') {
      dadosCI = r;
      Logger.log(`Dados CI encontrados: ${JSON.stringify(dadosCI)}`);
    }
  });
  
  // Preencher VOLUMES (J) - priorizar DI, depois CI
  let volumes = "";
  if (dadosDI?.QTDE_VOLUMES) {
    volumes = dadosDI.QTDE_VOLUMES;
  } else if (dadosCI?.QTDE_VOLUMES) {
    volumes = dadosCI.QTDE_VOLUMES;
    Logger.log(`VOLUMES do CI: ${volumes}`);
  }
  
  // Preencher PB (K) - priorizar DI, depois CI
  let pb = "";
  if (dadosDI?.PB) {
    pb = formatarParaExibicao(dadosDI.PB, '', 'PB');
  } else if (dadosCI?.PB) {
    pb = formatarParaExibicao(dadosCI.PB, '', 'PB');
    Logger.log(`PB do CI: ${pb}`);
  }
  
  // Preencher CIF (M) - priorizar CI, depois DI
  let cif = "";
  if (dadosCI?.CIF) {
    cif = formatarParaExibicao(dadosCI.CIF, 'R$ ', 'CIF');
    Logger.log(`CIF do CI: ${cif}`);
  } else if (dadosDI?.CIF) {
    cif = formatarParaExibicao(dadosDI.CIF, 'R$ ', 'CIF');
  }
  
  // Preencher as células
  if (volumes) {
    sheet.getRange(`J${linha}`).setValue(volumes);
    Logger.log(`✓ VOLUMES preenchido na célula J${linha}: ${volumes}`);
  }
  if (pb) {
    sheet.getRange(`K${linha}`).setValue(pb);
    Logger.log(`✓ PB preenchido na célula K${linha}: ${pb}`);
  }
  if (cif) {
    sheet.getRange(`M${linha}`).setValue(cif);
    Logger.log(`✓ CIF preenchido na célula M${linha}: ${cif}`);
  }
}

function formatarParaExibicao(valor, prefixo = '', campo = '') {
  if (!valor) return '';

  let numero;

  if (typeof valor === 'number') {
    numero = valor;
  } else {
    const valorStr = valor.toString();

  if (campo === 'CIF' || campo === 'VALORES_DEIM') {
    const valorNormalizado = valorStr.includes(',') 
      ? valorStr.replace(/\./g, '').replace(',', '.') 
      : valorStr;
      numero = parseFloat(valorNormalizado);
  } else if (campo === 'PB') {
      // Para PB, tratar como decimal direto
      const valorNumerico = parseFloat(valorStr.replace(',', '.'));
      return prefixo + valorNumerico.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    } else {
      // Para outros campos (II, IPI, TX_DI)
      const valorNormalizado = valorStr.includes(',') 
      ? valorStr.replace(/\./g, '').replace(',', '.') 
      : valorStr;
      numero = parseFloat(valorNormalizado);
    }

    if (isNaN(numero)) return prefixo + valor;
  }

  return prefixo + numero.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function formatarNumero(valor) {
  if (valor === null || valor === undefined || valor === "") return 0;
  const numero = parseFloat(String(valor).replace(/\./g, '').replace(',', '.'));
  return isNaN(numero) ? 0 : numero;
}

function extrairCIFCI(texto) {
  try {
    Logger.log("=== INICIANDO EXTRAÇÃO DADOS CI ===");
 
    const dados = {};
 
    // Procurar pela seção "DADOS SOBRE A CARGA" e extrair os valores em sequência
    const secaoCargaMatch = texto.match(/DADOS\s+SOBRE\s+A\s+CARGA[\s\S]*?VALOR\s+TOTAL\s+DA\s+IMPORTAÇÃO\s+\(R\$\)\s+PESO\s+BRUTO\s+\(Kg\)\s+QUANTIDADE\s+DE\s+VOLUMES\s+([\d.,]+)\s+([\d.,]+)\s+(\d+)/i);
 
    if (secaoCargaMatch) {
      // Os valores estão na ordem: CIF, PESO_BRUTO, QUANTIDADE_VOLUMES
      dados.CIF = secaoCargaMatch[1];
 
      // Processar o peso bruto mantendo formato brasileiro correto
      let valorPB = secaoCargaMatch[2];
      if (valorPB.includes(',')) {
        const numeroLimpo = valorPB.replace(/\./g, '').replace(',', '.');
        const numero = parseFloat(numeroLimpo);
        valorPB = numero.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
      }
      dados.PB = valorPB;
 
      dados.QTDE_VOLUMES = secaoCargaMatch[3];

      Logger.log(`✓ Dados extraídos da seção DADOS SOBRE A CARGA:`);
      Logger.log(`✓ CIF: "${dados.CIF}"`);
      Logger.log(`✓ PB: "${dados.PB}"`);
      Logger.log(`✓ QTDE_VOLUMES: "${dados.QTDE_VOLUMES}"`);
    } else {
      Logger.log("Seção DADOS SOBRE A CARGA não encontrada, tentando padrões individuais...");
 
      // Fallback para padrões individuais
      const qtdeVolumesMatch = texto.match(/QUANTIDADE\s+DE\s+VOLUMES\s+(\d+)/i) ||
                              texto.match(/Quantidade:\s*(\d+)/i);
      dados.QTDE_VOLUMES = qtdeVolumesMatch ? qtdeVolumesMatch[1] : "";
 
      const pesoBrutoMatch = texto.match(/PESO\s+BRUTO\s+\(Kg\)\s+([\d.,]+)/i) ||
                            texto.match(/PESO\s+BRUTO.*?([\d]+[,.][\d]+)/i);
      if (pesoBrutoMatch) {
        let valorPB = pesoBrutoMatch[1];
        if (valorPB.includes(',')) {
          const numeroLimpo = valorPB.replace(/\./g, '').replace(',', '.');
          const numero = parseFloat(numeroLimpo);
          valorPB = numero.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
        }
        dados.PB = valorPB;
      } else {
        dados.PB = "";
      }
 
      const cifMatch = texto.match(/VALOR\s+TOTAL\s+DA\s+IMPORTAÇÃO\s+\(R\$\)\s+([\d.,]+)/i) ||
                      texto.match(/VALOR\s+TOTAL\s+DA\s+IMPORTAÇÃO.*?([\d.,]+)/i);
      dados.CIF = cifMatch ? cifMatch[1] : "";
 
      Logger.log(`✓ Fallback - CIF: "${dados.CIF}", PB: "${dados.PB}", QTDE_VOLUMES: "${dados.QTDE_VOLUMES}"`);
}
 
    Logger.log(`=== DADOS CI EXTRAÍDOS: ${JSON.stringify(dados)} ===`);
    return dados;
  } catch (e) {
    Logger.log(`Erro ao extrair dados do CI: ${e}`);
    return null;
  }
}

// Função para extrair dados específicos do documento DI
function extrairDadosDI(texto) {
  const resultado = {};
  
  try {
    // Extrair número da DI
    const regexNumeroDI = /(?:di\s*n[oº°]?|declaração\s*de\s*importação)[\s\:\-]*(\d{11})/i;
    const matchDI = texto.match(regexNumeroDI);
    if (matchDI) {
      resultado.numeroDI = matchDI[1];
      Logger.log(`Número DI extraído: ${resultado.numeroDI}`);
    }
    
    // Extrair data de registro
    const regexData = /(?:data\s*de\s*registro|reg\.\s*em)[\s\:\-]*(\d{2}\/\d{2}\/\d{4})/i;
    const matchData = texto.match(regexData);
    if (matchData) {
      resultado.dataRegistro = matchData[1];
      Logger.log(`Data de registro extraída: ${resultado.dataRegistro}`);
    }
    
    // Extrair valor aduaneiro
    const regexValorAduaneiro = /(?:valor\s*aduaneiro|valor\s*total)[\s\:\-]*(?:usd?|us\$|\$)?\s*(\d+(?:\.\d+)?(?:,\d+)?)/i;
    const matchValor = texto.match(regexValorAduaneiro);
    if (matchValor) {
      resultado.valorAduaneiro = parseFloat(matchValor[1].replace(',', '.'));
      Logger.log(`Valor aduaneiro extraído: ${resultado.valorAduaneiro}`);
    }
    
    Logger.log(`Dados DI extraídos: ${JSON.stringify(resultado)}`);
    return Object.keys(resultado).length > 0 ? resultado : null;
    
  } catch (error) {
    Logger.log(`Erro ao extrair dados DI: ${error}`);
    return null;
  }
}

function findSubpasta(pastaRaiz, idInserido) {
  const cacheKey = `pasta_${idInserido}`;
  
  // Verificar cache primeiro
  if (CACHE_MPR.pastasCache[cacheKey]) {
    try {
      // Verificar se a pasta ainda existe
      const pastaCache = DriveApp.getFolderById(CACHE_MPR.pastasCache[cacheKey]);
      Logger.log(`✓ Usando pasta do cache: ${pastaCache.getName()}`);
      return pastaCache;
    } catch (error) {
      // Pasta não existe mais, remover do cache
      delete CACHE_MPR.pastasCache[cacheKey];
      Logger.log(`Cache de pasta inválido removido para: ${idInserido}`);
    }
  }

  Logger.log(`🔍 Buscando subpasta para: ${idInserido}`);
  
  try {
    const pastasEncontradas = [];
    
    // Listar todas as pastas mensais primeiro
    const pastasMensais = pastaRaiz.getFolders();
    Logger.log("📁 Pastas mensais disponíveis:");
    
    while (pastasMensais.hasNext()) {
      const pastaMensal = pastasMensais.next();
      const nomePastaMensal = pastaMensal.getName();
      Logger.log(`   📂 ${nomePastaMensal}`);
      
      // Buscar dentro de cada pasta mensal
      const subpastas = pastaMensal.getFolders();
      
      while (subpastas.hasNext()) {
        const subpasta = subpastas.next();
        const nomeSubpasta = subpasta.getName();
        
        // Primeira tentativa: busca exata pelo padrão esperado
        if (nomeSubpasta.endsWith(`_${idInserido}`) && 
            /^P\d+(-25)?_/.test(nomeSubpasta)) {
          
          Logger.log(`✓ Subpasta encontrada (padrão exato) em ${nomePastaMensal}: ${nomeSubpasta}`);
          CACHE_MPR.pastasCache[cacheKey] = subpasta.getId();
          return subpasta;
        }
        
        // Segunda tentativa: busca flexível - pasta que contém o ID
        if (nomeSubpasta.includes(idInserido)) {
          pastasEncontradas.push({
            pasta: subpasta,
            nome: nomeSubpasta,
            pastaMensal: nomePastaMensal,
            score: calcularScoreSimilaridade(nomeSubpasta, idInserido)
          });
          Logger.log(`   🔍 Pasta candidata em ${nomePastaMensal}: ${nomeSubpasta}`);
        }
      }
    }
    
    // Se encontrou pastas que contêm o ID, escolher a melhor
    if (pastasEncontradas.length > 0) {
      // Ordenar por score de similaridade (maior primeiro)
      pastasEncontradas.sort((a, b) => b.score - a.score);
      
      const melhorPasta = pastasEncontradas[0];
      Logger.log(`✓ Subpasta encontrada (busca flexível) em ${melhorPasta.pastaMensal}: ${melhorPasta.nome} (score: ${melhorPasta.score})`);
      
      if (pastasEncontradas.length > 1) {
        Logger.log(`⚠️ Múltiplas pastas encontradas. Outras opções:`);
        pastasEncontradas.slice(1).forEach(pasta => {
          Logger.log(`   - ${pasta.nome} em ${pasta.pastaMensal} (score: ${pasta.score})`);
        });
      }
      
      // Armazenar no cache
      CACHE_MPR.pastasCache[cacheKey] = melhorPasta.pasta.getId();
      return melhorPasta.pasta;
    }
    
    Logger.log(`❌ Subpasta não encontrada para: ${idInserido}`);
    return null;
    
  } catch (error) {
    Logger.log(`❌ Erro ao buscar subpasta: ${error.message}`);
    return null;
  }
}


// Função auxiliar para calcular similaridade entre nomes
function calcularScoreSimilaridade(nomePasta, idInserido) {
  let score = 0;
  
  // Score básico se contém o ID
  if (nomePasta.includes(idInserido)) {
    score += 10;
  }
  
  // Score adicional se termina com o ID (padrão esperado)
  if (nomePasta.endsWith(`_${idInserido}`)) {
    score += 20;
  }
  
  // Score adicional se segue o padrão P123456-25_
  if (/^P\d+(-25)?_/.test(nomePasta)) {
    score += 15;
  }
  
  // Score adicional se tem exatamente o padrão completo
  if (/^P\d{6}-25_/.test(nomePasta) && nomePasta.endsWith(`_${idInserido}`)) {
    score += 25;
  }
  
  // Penalizar se o nome é muito longo (pode ser falso positivo)
  if (nomePasta.length > 50) {
    score -= 5;
  }
  
  return score;
}

// Função para limpar todo o cache (útil para debug)
function limparTodoCache() {
  CACHE_MPR.planilhaPrimaria = null;
  CACHE_MPR.mapeamentoIndices = null;
  CACHE_MPR.timestampCache = null;
  CACHE_MPR.pastasCache = {};
  CACHE_MPR.arquivosProcessados = {};
  CACHE_MPR.ultimaLimpeza = null;
  
  Logger.log("🧹 Todo o cache foi limpo");
}

// Função para testar o sistema (útil para debug)
function testarSistema() {
  try {
    Logger.log("🧪 Iniciando teste do sistema...");
    
    // Testar pré-carregamento
    const resultado = preCarregarPlanilhaPrimaria(true);
    Logger.log(`✓ Pré-carregamento OK: ${resultado.dados.length - 1} registros`);
    
    // Verificar status do cache
    verificarStatusCache();
    
    Logger.log("✅ Sistema funcionando corretamente");
    return true;
    
  } catch (error) {
    Logger.log(`❌ Erro no teste: ${error}`);
    return false;
  }
}

// =============================================================================================================//
// FUNÇÕES DE EMAIL

function obterDestinatarios() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Config");
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const nomeAba = sheet.getName();
  
  if (!configSheet) {
    Logger.log("Aviso: Planilha 'Config' não encontrada. Usando e-mails padrão.");
    return {
      to: ["logistica@empresa.com"],
      cc: [],
      bcc: []
    };
  }

  const data = configSheet.getDataRange().getValues();
  const destinatarios = { to: [], cc: [], bcc: [] };

  // Estrutura esperada da planilha Config:
  // Coluna A: Tipo (GERAL, VCP, GRU, CWB, PNG MARÍTIMO)
  // Coluna B: Categoria (principal, cópia, cco)
  // Coluna C: Email

  for (let i = 1; i < data.length; i++) {
    const tipoRecinto = data[i][0] ? data[i][0].toString().toUpperCase().trim() : '';
    const categoria = data[i][1] ? data[i][1].toString().toLowerCase().trim() : '';
    const email = data[i][2] ? data[i][2].toString().trim() : '';
    
    if (!email) continue;

    // Verificar se é para este recinto específico ou geral
    const isParaEsteRecinto = tipoRecinto === nomeAba.toUpperCase() || tipoRecinto === 'GERAL';
    
    if (!isParaEsteRecinto) continue;

    if (categoria.includes('principal')) {
      destinatarios.to.push(email);
    } 
    else if (categoria.includes('cópia') || categoria.includes('copia')) {
      destinatarios.cc.push(email);
    } 
    else if (categoria.includes('cco') || categoria.includes('oculta')) {
      destinatarios.bcc.push(email);
    } 
    else {
      destinatarios.cc.push(email);
    }
  }

  // Se não encontrou destinatários principais, usar padrão
  if (destinatarios.to.length === 0) {
    destinatarios.to.push("logistica@empresa.com");
  }

  Logger.log(`Destinatários para ${nomeAba}: TO: ${destinatarios.to.join(', ')}, CC: ${destinatarios.cc.join(', ')}, BCC: ${destinatarios.bcc.join(', ')}`);
  
  return destinatarios;
}

function determinarTipoProcesso(nomeAba, dadosProcesso) {
  // Mapear diretamente pela aba
  switch (nomeAba.toUpperCase()) {
    case "VCP":
      return "VCP";
    case "GRU":
      return "GRU";
    case "CWB":
      return "CWB";
    case "PNG MARÍTIMO":
      return "PNG_MARITIMO";
    default:
      Logger.log(`Aba não reconhecida: ${nomeAba}`);
      return "DESCONHECIDO";
  }
}

function coletarAnexos(linhas, cabecalhos) {
  const anexos = [];
  const pastaRaiz = DriveApp.getFolderById("1NjYLk2dt8PB3RpaJ0C9O8K_Gs_8OIHoZ");
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const nomeAba = sheet.getName();
  
  // Determinar tipo de processo pela aba
  const tipoProcesso = determinarTipoProcesso(nomeAba, {});
  
  Logger.log(`Coletando anexos para processo tipo: ${tipoProcesso}`);
  
  linhas.forEach(linha => {
    const id = linha[0].toString().trim();
    if (!id) return;
    
    Logger.log(`Coletando anexos para ID: ${id}`);
    
    const subpasta = findSubpasta(pastaRaiz, id);
    if (!subpasta) {
      Logger.log(`Subpasta não encontrada para ID: ${id}`);
      return;
    }
    
    const arquivos = subpasta.getFiles();
    const idUpper = id.toUpperCase();
    
    while (arquivos.hasNext()) {
      const arquivo = arquivos.next();
      const nome = arquivo.getName().toUpperCase();
      
      let deveAnexar = false;
      let tipoAnexo = "";
      
      // Definir quais anexos coletar baseado no tipo de processo
      switch (tipoProcesso) {
        case "VCP":
          if (nome.includes(`_DI_${idUpper}.PDF`)) {
            deveAnexar = true;
            tipoAnexo = "DI";
          }
          break;
        case "GRU":
          
          if (nome.includes(`_CI_${idUpper}.PDF`)) {
            deveAnexar = true;
            tipoAnexo = "CI";
          } else if (nome.includes(`_DI_${idUpper}.PDF`)) {
            deveAnexar = true;
            tipoAnexo = "DI";
          }
          break;
          
        case "CWB":
          // CWB: CI, DI, CCT e GLME
          if (nome.includes(`_CI_${idUpper}.PDF`)) {
            deveAnexar = true;
            tipoAnexo = "CI";
          } else if (nome.includes(`_DI_${idUpper}.PDF`)) {
            deveAnexar = true;
            tipoAnexo = "DI";
          } else if (nome.includes(`_CCT_${idUpper}.PDF`) || (nome.includes(`CCT`) && nome.includes(idUpper))) {
            deveAnexar = true;
            tipoAnexo = "CCT";
          } else if (nome.includes(`_GLME_${idUpper}.PDF`) || (nome.includes(`GLME`) && nome.includes(idUpper))) {
            deveAnexar = true;
            tipoAnexo = "GLME";
          }
          break;
          
        case "PNG_MARITIMO":
          // PNG MARÍTIMO: CI, DI e GLME
          if (nome.includes(`_CI_${idUpper}.PDF`)) {
            deveAnexar = true;
            tipoAnexo = "CI";
          } else if (nome.includes(`_DI_${idUpper}.PDF`)) {
            deveAnexar = true;
            tipoAnexo = "DI";
          } else if (nome.includes(`_GLME_${idUpper}.PDF`) || (nome.includes(`GLME`) && nome.includes(idUpper))) {
            deveAnexar = true;
            tipoAnexo = "GLME";
          }
          break;
      }
      
      if (deveAnexar) {
        Logger.log(`Anexo ${tipoProcesso} encontrado: ${arquivo.getName()} (${tipoAnexo})`);
        anexos.push({
          arquivo: arquivo,
          nome: arquivo.getName(),
          id: id,
          tipo: tipoAnexo,
          tipoProcesso: tipoProcesso
        });
      }
    }
  });
  
  Logger.log(`Total de anexos coletados: ${anexos.length}`);
  return anexos;
}

function criarCorpoEmail(tipoProcesso, dataFormatada, tabelaHTML) {
  const agora = new Date();
  const saudacao = agora.getHours() < 12 ? "Bom dia" : "Boa tarde";
  
  let corpo = "";
  
  switch (tipoProcesso) {
    case "VCP":
      corpo = `
        Prezados, ${saudacao.toLowerCase()}.<br><br>
        Segue(m) registro(s) Renault - VCP.<br>
        <br>
        <span style="background-color: yellow; font-weight: bold;">Em anexo DI.</span><br><br>
        ${tabelaHTML}<br><br>
      `;
      break;
      
    case "GRU":
      corpo = `
        Prezados, ${saudacao.toLowerCase()}.<br><br>
        Segue(m) registro(s) Renault - GRU.<br>
        <br>
        <span style="background-color: yellow; font-weight: bold;">Em anexo CI e DI.</span><br><br>
        ${tabelaHTML}<br><br>
      `;
      break;
      
    case "CWB":
      corpo = `
        ${saudacao},<br><br>
        Segue(m) registro(s) Renault - CWB.<br>
        <br>
        <span style="background-color: yellow; font-weight: bold;">Em anexo CI, DI, GLME e extrato do CCT.</span><br><br>
        ${tabelaHTML}<br><br>
      `;
      break;
      
    case "PNG_MARITIMO":
      corpo = `
        ${saudacao},<br><br>
        Segue(m) registro(s) MPR RENAULT.<br>
        <br>
        <span style="background-color: yellow; font-weight: bold;">Em anexo CI, DI e GLME.</span><br><br>
        Navio: <br>
        ATA: <br><br>
        ${tabelaHTML}<br><br>
        
      `;
      break;
      
    default:
      corpo = `
        ${saudacao},<br><br>
        Segue(m) registros Renault.<br>
        <br>
        ${tabelaHTML}<br><br>
      `;
  }
  
  return corpo;
}

function coletarFaturas(linhas, cabecalhos) {
  const faturas = [];
  const indiceFatura = cabecalhos.indexOf("FATURA");
  
  if (indiceFatura === -1) {
    Logger.log("Coluna FATURA não encontrada");
    return faturas;
  }
  
  linhas.forEach(linha => {
    const fatura = linha[indiceFatura];
    if (fatura && fatura.toString().trim() !== "") {
      const faturaStr = fatura.toString().trim();
      if (!faturas.includes(faturaStr)) {
        faturas.push(faturaStr);
      }
    }
  });
  
  return faturas;
}

function criarAssuntoEmail(faturas, nomeAba) {
  const faturasStr = faturas.length > 0 ? faturas.join(" ") : "";
  return `MPR - REGISTRO DI - ${nomeAba} - ${faturasStr} - CNPJ 001900`;
}

function criarRascunhoEmailComAnexos() {
  const CONFIG = {
    timeZone: Session.getScriptTimeZone(),
    dateFormat: "dd/MM",
    maxLinhas: 15,
    estiloCelula: "border: 1px solid #000; padding: 4px; vertical-align: middle; text-align: center;",
    cores: {
      cabecalho: "#f8dc75",
      fonte: "Calibri"
    }
  };

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const nomeAba = sheet.getName();
  
  // Verificar se é uma aba válida
  const abasValidas = ["VCP", "GRU", "CWB", "PNG MARÍTIMO"];
  if (!abasValidas.includes(nomeAba)) {
    Logger.log(`Aba ${nomeAba} não configurada para envio de email.`);
    SpreadsheetApp.getUi().alert(`Aba "${nomeAba}" não está configurada para envio de email.\nAbas válidas: ${abasValidas.join(', ')}`);
    return;
  }

  const data = sheet.getDataRange().getValues();
  const cabecalhos = data[0];
  const linhas = data.slice(1, CONFIG.maxLinhas + 1).filter(row => row[0] !== "");

  if (linhas.length === 0) {
    Logger.log("Nenhum dado encontrado para processar.");
    SpreadsheetApp.getUi().alert("Nenhum dado encontrado para processar.");
    return;
  }

  const destinatarios = obterDestinatarios();
  const agora = new Date();
  const dataFormatada = Utilities.formatDate(agora, CONFIG.timeZone, CONFIG.dateFormat);

  // Encontrar o índice da coluna CIF para limitar até ela
  const indiceCIF = cabecalhos.indexOf("CIF");
  const colunasFiltradas = indiceCIF !== -1 ? cabecalhos.slice(0, indiceCIF + 1) : cabecalhos;
  
  Logger.log(`Exibindo colunas até CIF. Total de colunas: ${colunasFiltradas.length}`);

  // Determinar larguras das colunas baseado no tipo de aba
  let largurasColunas;
  if (nomeAba === "VCP" || nomeAba === "GRU" || nomeAba === "CWB") {
    largurasColunas = [
      "140px", "70px", "110px", "130px", "150px", "200px",
      "110px", "120px", "80px", "80px", "100px", "120px"
    ].slice(0, colunasFiltradas.length);
  } else if (nomeAba === "PNG MARÍTIMO") {
    largurasColunas = [
      "140px", "100px", "110px", "130px", "150px", "120px", 
      "130px", "140px", "80px", "100px", "80px", "120px"
    ].slice(0, colunasFiltradas.length);
  } else {
    largurasColunas = Array(colunasFiltradas.length).fill("100px");
  }

  const formatarNumeroBrasileiro = (valor) => {
    if (typeof valor === "string") {
      valor = parseFloat(valor.replace(",", "."));
    }
    return typeof valor === "number" && !isNaN(valor) ? 
      valor.toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : 
      valor;
  };

  const construirTabelaHTML = () => {
    let html = `<table style="border-collapse: collapse; font-family: ${CONFIG.cores.fonte};">`;
    
    // Cabeçalho da tabela
    html += '<tr>' + colunasFiltradas.map((cabecalho, index) => 
      `<th style="${CONFIG.estiloCelula} background-color: ${CONFIG.cores.cabecalho}; width: ${largurasColunas[index] || '100px'};">${cabecalho}</th>`
    ).join('') + '</tr>';
    
    // Linhas da tabela
    linhas.forEach(linha => {
      const linhaLimitada = linha.slice(0, colunasFiltradas.length);
      html += '<tr>' + linhaLimitada.map((celula, index) => {
        const nomeCabecalho = colunasFiltradas[index];
        let valorFormatado = celula;
        
        if (nomeCabecalho === "CIF") {
          valorFormatado = `R$ ${formatarNumeroBrasileiro(celula)}`;
        } else if (nomeCabecalho === "TOTAL") {
          valorFormatado = `R$ ${formatarNumeroBrasileiro(celula)}`;
        } else if (["PB", "II", "IPI", "TX_DI", "VALORES_DEIM"].includes(nomeCabecalho)) {
          valorFormatado = formatarNumeroBrasileiro(celula);
        }
        
        return `<td style="${CONFIG.estiloCelula} width: ${largurasColunas[index] || '100px'};">${valorFormatado}</td>`;
      }).join('') + '</tr>';
    });
    
    html += '</table>';
    return html;
  };

  // Determinar o tipo de processo baseado na aba
  const tipoProcesso = determinarTipoProcesso(nomeAba, {});
  
  Logger.log(`Tipo de processo detectado: ${tipoProcesso}`);

  // Coletar anexos baseado no tipo de processo
  const anexosColetados = coletarAnexos(linhas, cabecalhos);
  const anexosBlob = anexosColetados.map(anexo => {
    try {
      return anexo.arquivo.getBlob();
    } catch (error) {
      Logger.log(`Erro ao processar anexo ${anexo.nome}: ${error}`);
      return null;
    }
  }).filter(blob => blob !== null);

  // Construir tabela HTML
  const tabelaHTML = construirTabelaHTML();
  
  // Criar corpo do email baseado no tipo de processo
  const corpoEmail = criarCorpoEmail(tipoProcesso, dataFormatada, tabelaHTML);
  
  // Coletar todas as faturas para o assunto
  const faturas = coletarFaturas(linhas, cabecalhos);
  
  // Criar assunto do email com todas as faturas
  const assuntoEmail = criarAssuntoEmail(faturas, nomeAba);

  const opcoes = {
    htmlBody: corpoEmail,
    cc: destinatarios.cc.join(','),
    bcc: destinatarios.bcc.join(',')
  };

  if (anexosBlob.length > 0) {
    opcoes.attachments = anexosBlob;
    Logger.log(`Anexando ${anexosBlob.length} arquivo(s) ao rascunho`);
  }

  try {
    GmailApp.createDraft(
      destinatarios.to.join(','),
      assuntoEmail,
      "",
      opcoes
    );

    Logger.log(`Rascunho criado para: ${destinatarios.to.join(', ')} com assunto: ${assuntoEmail}`);
    Logger.log(`Tipo de processo: ${tipoProcesso}`);
    Logger.log(`Faturas incluídas: ${faturas.join(', ')}`);
    Logger.log(`Anexos incluídos: ${anexosColetados.map(a => `${a.nome} (${a.tipo})`).join(', ')}`);
    
    SpreadsheetApp.getUi().alert(`Rascunho criado com sucesso!\n\nDestinatários: ${destinatarios.to.join(', ')}\nAnexos: ${anexosColetados.length}\nTipo: ${tipoProcesso}`);
    
  } catch (error) {
    Logger.log(`Erro ao criar rascunho: ${error}`);
    SpreadsheetApp.getUi().alert(`Erro ao criar rascunho: ${error}`);
  }
}

function listarAnexosDisponiveis() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const linhas = data.slice(1, 16).filter(row => row[0] !== "");
  
  Logger.log("=== LISTAGEM DE ANEXOS DISPONÍVEIS ===");
  
  const anexos = coletarAnexos(linhas, data[0]);
  
  anexos.forEach(anexo => {
    Logger.log(`ID: ${anexo.id} | Processo: ${anexo.tipoProcesso} | Tipo: ${anexo.tipo} | Arquivo: ${anexo.nome}`);
  });
  
  Logger.log(`Total encontrado: ${anexos.length} anexos`);
}


function testarColeta() {
  try {
    listarAnexosDisponiveis();
  } catch (error) {
    Logger.log(`Erro no teste: ${error}`);
  }
}

function onOpen() {
  adicionarMenus();
}

function adicionarMenus() {
  const ui = SpreadsheetApp.getUi();
  
  // Menu único que funciona para ambas as abas
  ui.createMenu('📧 E-mail Rascunho')
    .addItem('Criar Rascunho de Email', 'criarRascunhoEmailComAnexos')
    .addItem('Listar Anexos Disponíveis', 'listarAnexosDisponiveis')
    .addToUi();
}