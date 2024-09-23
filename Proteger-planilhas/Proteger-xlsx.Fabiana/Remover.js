function removerProtecoesEmAbasEspecificas() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var nomesAbas = [
    "BP JUNDIAÍ",
    "BP ALPHAVILLE",
    "SP MOEMA",
    "BP RONDÔNIA",
    "BP RIO GRANDE DO NORTE",
    "BP SUDESTE PAULISTA",
  ];

  // Para cada aba especificada
  nomesAbas.forEach(function (nomeAba) {
    var aba = planilha.getSheetByName(nomeAba);
    if (aba) {
      var protecoes = aba.getProtections(SpreadsheetApp.ProtectionType.RANGE);

      // Para cada proteção na aba, removê-la
      protecoes.forEach(function (protecao) {
        protecao.remove();
      });
    }
  });
}
