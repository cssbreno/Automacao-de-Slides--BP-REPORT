function removerProtecoesEmAbasEspecificas() {
    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var nomesAbas = ['BP CURITIBA', 'BP NORTE CATARINENSE', 'BP SANTA CATARINA', 'BP LITORAL CATARINENSE', 'BP VITÓRIA DA CONQUISTA', 'DC TM'];
  
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
  