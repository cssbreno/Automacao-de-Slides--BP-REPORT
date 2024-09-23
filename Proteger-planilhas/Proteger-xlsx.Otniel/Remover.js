function removerProtecoesEmAbasEspecificas() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var nomesAbas = [
    "BP SP BERRINI",
    "BP MATO GROSSO DO SUL",
    "BP MATO GROSSO",
    "BP ALAGOAS",
    "DC SÃO PAULO",
    "BP NORTE DE MINAS",
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
