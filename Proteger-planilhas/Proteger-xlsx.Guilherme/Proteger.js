function protegerCelulasEmAbasEspecificas() {
    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var nomesAbas = ['BP SP METROPOLITANA', 'BP INDAIATUBA', 'BP MARANHÃO', 'BP PARÁ', 'BP RECIFE', 'DC CENTRO SUL', 'BP ABC PAULISTA'];
    var intervalosProtegidos = [
      '148:148', '139:139', '193:193', '202:202', 'B:B', 'A:A', 
      '74:74', 'J:L', '75:75', '117:117', 'D25:D31', '106:106', 
      '167:167', '149:149', '102:102', 'D13:D19', '91:91', 
      '194:194', '118:118', '212:212', '107:107', 
      '90:90', '48:48', 'D1:D7', '168:168', '130:132', 
      'G:G', '185:187', '58:58', '13:13', '25:25', 
      '211:211', '138:138', '37:37', '59:59', '1:1', 
      '159:159', '158:158', '47:47', '203:203'
    ];
  
    // Para cada aba especificada
    nomesAbas.forEach(function(nomeAba) {
      var aba = planilha.getSheetByName(nomeAba);
      if (aba) {
        intervalosProtegidos.forEach(function(intervalo) {
          var range = aba.getRange(intervalo);
          var protecao = range.protect().setDescription('Proteção automática');
          
          // Permite edição apenas para o seu e-mail
          protecao.removeEditors(protecao.getEditors());
          protecao.addEditor('breno.soares@sankhya.com.br'); // Adiciona seu e-mail
  
          // (Opcional) Caso queira permitir edição para o proprietário
          if (protecao.canDomainEdit()) {
            protecao.setDomainEdit(false);
          }
        });
      }
    });
  }
  