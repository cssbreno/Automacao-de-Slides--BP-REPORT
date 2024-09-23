function removerProtecoesEmAbasEspecificas() {
    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var nomesAbas = ['BP SALVADOR', 'DC SUDESTE', 'BP CEARÁ', 'BP BH PAMPULHA', 'BP SUL DE MINAS', 'BP RJ'];
  
    // Para cada aba especificada
    nomesAbas.forEach(function(nomeAba) {
      var aba = planilha.getSheetByName(nomeAba);
      if (aba) {
        var protecoes = aba.getProtections(SpreadsheetApp.ProtectionType.RANGE);
        
        // Para cada proteção na aba, removê-la
        protecoes.forEach(function(protecao) {
          protecao.remove();
        });
      }
    });
  }
  