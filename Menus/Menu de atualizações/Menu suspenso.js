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
    .setHeight(440);
  SlidesApp.getUi().showDialog(htmlOutput);
}

function showRomulloMenu() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('RomulloMenu')
    .setWidth(350)
    .setHeight(440);
  SlidesApp.getUi().showDialog(htmlOutput);
}

function showOtnielMenu() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('OtnielMenu')
    .setWidth(350)
    .setHeight(440);
  SlidesApp.getUi().showDialog(htmlOutput);
}

function showGuilhermeMenu() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('GuilhermeMenu')
    .setWidth(350)
    .setHeight(480);
  SlidesApp.getUi().showDialog(htmlOutput);
}

function showFabianaMenu() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('FabianaMenu')
    .setWidth(350)
    .setHeight(440);
  SlidesApp.getUi().showDialog(htmlOutput);
}

// ------------------------ FUNÇÕES DA VANESSA ----------------------------------------------

function updateBP_SALVADOR() {
  Logger.log("Iniciando BP SALVADOR");
  BPsVanessaGozdziuk.updateBP_SALVADOR();
}

function updateDC_SUDESTE() {
  Logger.log("Iniciando DC SUDESTE");
  BPsVanessaGozdziuk.updateDC_SUDESTE();
}

function updateBP_CE() {
  Logger.log("Iniciando BP CE");
  BPsVanessaGozdziuk.updateBP_CE();
}

function updateBP_BHPAMPULHA() {
  Logger.log("Iniciando BP BH PAMPULHA");
  BPsVanessaGozdziuk.updateBP_BHPAMPULHA();
}

function updateBP_SULMG() {
  Logger.log("Iniciando BP SUL DE MINAS");
  BPsVanessaGozdziuk.updateBP_SULMG();
}

function updateBP_RJ() {
  Logger.log("Iniciando BP RJ");
  BPsVanessaGozdziuk.updateBP_RJ();
}

// ------------------------ FUNÇÕES DO ROMULLO ----------------------------------------------

function updateBP_CWB() {
  Logger.log("Iniciando BP CWB");
  BPsRomulloCarvalho.updateBP_CWB();
}

function updateBP_NCATARINENSE() {
  Logger.log("Iniciando BP NCATARINENSE");
  BPsRomulloCarvalho.updateBP_NCATARINENSE();
}

function updateBP_SC() {
  Logger.log("Iniciando BP SC");
  BPsRomulloCarvalho.updateBP_SC();
}

function updateBP_LITCAT() {
  Logger.log("Iniciando BP LITCAT");
  BPsRomulloCarvalho.updateBP_LITCAT();
}

function updateBP_VCONQ() {
  Logger.log("Iniciando BP VCONQ");
  BPsRomulloCarvalho.updateBP_VCONQ();
}

function updateDC_TM() {
  Logger.log("Iniciando DC TM");
  BPsRomulloCarvalho.updateDC_TM();
}

// ------------------------ FUNÇÕES DO OTNIEL ----------------------------------------------

function updateBP_BERRINI() {
  Logger.log("Iniciando BP BERRINI");
  BPsOtnieldeFarias.updateBP_BERRINI();
}

function updateBP_MS() {
  Logger.log("Iniciando BP MS");
  BPsOtnieldeFarias.updateBP_MS();
}

function updateBP_MT() {
  Logger.log("Iniciando BP MT");
  BPsOtnieldeFarias.updateBP_MT();
}

function updateBP_ALAGOAS() {
  Logger.log("Iniciando BP ALAGOAS");
  BPsOtnieldeFarias.updateBP_ALAGOAS();
}

function updateDC_SP() {
  Logger.log("Iniciando DC SP");
  BPsOtnieldeFarias.updateDC_SP();
}

function updateBP_NORTEMG() {
  Logger.log("Iniciando BP NORTEMG");
  BPsOtnieldeFarias.updateBP_NORTEMG();
}

// ------------------------ FUNÇÕES DO GUILHERME ----------------------------------------------

function updateBP_SPMET() {
  Logger.log("Iniciando BP SPMET");
  BPsGuilhermeMedeiros.updateBP_SPMET();
}

function updateBP_INDAIATUBA() {
  Logger.log("Iniciando BP INDAIATUBA");
  BPsGuilhermeMedeiros.updateBP_INDAIATUBA();
}

function updateBP_MA() {
  Logger.log("Iniciando BP MARANHÃO");
  BPsGuilhermeMedeiros.updateBP_MA();
}

function updateBP_PA() {
  Logger.log("Iniciando BP PARÁ");
  BPsGuilhermeMedeiros.updateBP_PA();
}

function updateBP_RECIFE() {
  Logger.log("Iniciando BP RECIFE");
  BPsGuilhermeMedeiros.updateBP_RECIFE();
}

function updateDC_CSUL() {
  Logger.log("Iniciando DC CENTRO SUL");
  BPsGuilhermeMedeiros.updateDC_CSUL();
}

function updateBP_ABCP() {
  Logger.log("Iniciando BP ABC PAULISTA");
  BPsGuilhermeMedeiros.updateBP_ABCP();
}

// ------------------------ FUNÇÕES DA FABIANA  ----------------------------------------------

function updateBP_JUNDIAI() {
  Logger.log("Iniciando BP JUNDIAI");
  BPsFabianaVerissimo.updateBP_JUNDIAI();
}

function updateBP_ALPHAVILLE() {
  Logger.log("Iniciando BP ALPHAVILLE");
  BPsFabianaVerissimo.updateBP_ALPHAVILLE();
}

function updateBP_MOEMA() {
  Logger.log("Iniciando BP MOEMA");
  BPsFabianaVerissimo.updateBP_MOEMA();
}

function updateBP_RONDONIA() {
  Logger.log("Iniciando BP RONDÔNIA");
  BPsFabianaVerissimo.updateBP_RONDONIA();
}

function updateBP_RN() {
  Logger.log("Iniciando BP RIO GRANDE DO NORTE");
  BPsFabianaVerissimo.updateBP_RN();
}

function updateBP_SUDESTEP() {
  Logger.log("Iniciando BP SUDESTE PAULISTA");
  BPsFabianaVerissimo.updateBP_SUDESTEP();
}

function processSelection(option) {
  Logger.log('Você escolheu: ' + option);
}
