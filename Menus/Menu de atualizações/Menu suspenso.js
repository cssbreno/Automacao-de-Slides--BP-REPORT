function onOpen() {
    var ui = SlidesApp.getUi();
    ui.createMenu('Menu BPS')
      .addItem('BPS - Vanessa', 'showVanessaMenu')
      .addItem('BPS - Romullo', 'showRomulloMenu')
      .addItem('BPS - Otniel', 'showOtnielMenu')
      .addItem('BPS - Guilherme', 'showGuilhermeMenu')
      .addItem('BPS - Fabiana', 'showFabianaMenu')
      .addToUi();
  }
  
  // Funções para abrir diálogos com opções de submenu
  function showVanessaMenu() {
    var htmlOutput = HtmlService.createHtmlOutputFromFile('VanessaMenu')
        .setWidth(350)
        .setHeight(400);
    SlidesApp.getUi().showDialog(htmlOutput);
  }
  
  function showRomulloMenu() {
    var htmlOutput = HtmlService.createHtmlOutputFromFile('RomulloMenu')
        .setWidth(350)
        .setHeight(400);
    SlidesApp.getUi().showDialog(htmlOutput);
  }
  
  function showOtnielMenu() {
    var htmlOutput = HtmlService.createHtmlOutputFromFile('OtnielMenu')
        .setWidth(350)
        .setHeight(400);
    SlidesApp.getUi().showDialog(htmlOutput);
  }
  
  function showGuilhermeMenu() {
    var htmlOutput = HtmlService.createHtmlOutputFromFile('GuilhermeMenu')
        .setWidth(350)
        .setHeight(400);
    SlidesApp.getUi().showDialog(htmlOutput);
  }
  
  function showFabianaMenu() {
    var htmlOutput = HtmlService.createHtmlOutputFromFile('FabianaMenu')
        .setWidth(350)
        .setHeight(400);
    SlidesApp.getUi().showDialog(htmlOutput);
  }
  
  // Funções para cada opção do submenu de Vanessa usando a biblioteca
  function updateBP_SALVADOR() {
    BPsVanessaGozdziuk.updateBP_SALVADOR(); // Chamada direta da função da biblioteca
  }
  
  function updateBP_SUDESTE() {
    BPsVanessaGozdziuk.updateBP_SUDESTE(); // Chamada direta da função da biblioteca
  }
  
  function updateBP_CEARA() {
    BPsVanessaGozdziuk.updateBP_CEARA(); // Chamada direta da função da biblioteca
  }
  
  function updateBP_BH_PAMPULHA() {
    BPsVanessaGozdziuk.updateBP_BH_PAMPULHA(); // Chamada direta da função da biblioteca
  }
  
  function updateBP_SUL_DE_MINAS() {
    BPsVanessaGozdziuk.updateBP_SUL_DE_MINAS(); // Chamada direta da função da biblioteca
  }
  
  function updateBP_RJ() {
    BPsVanessaGozdziuk.updateBP_RJ(); // Chamada direta da função da biblioteca
  }
  
  // Função para processar a seleção de submenu (se necessário)
  function processSelection(option) {
    Logger.log('Você escolheu: ' + option);
    // Adicione a lógica para tratar a opção selecionada aqui
  }
  