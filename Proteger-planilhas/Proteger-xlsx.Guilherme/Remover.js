function removerProtecoesEmAbasEspecificas() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var nomesAbas = [
    "BP SP METROPOLITANA",
    "BP INDAIATUBA",
    "BP MARANHÃO",
    "BP PARÁ",
    "BP RECIFE",
    "DC CENTRO SUL",
    "BP ABC PAULISTA",
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
