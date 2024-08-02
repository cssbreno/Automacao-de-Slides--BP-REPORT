// Atualizar informações da planilha para o slide
function updateSlidesFromSheet() {

  // IDs da apresentação e da planilha
  const SLIDES_ID = '1whKu0ncTiW3qFmPUXusqTL7zu_ai_uy4kz09oyoaq34';
  const SHEETS_ID = '1GZAPqryb7C3J_rF2GpyMAYkqj5itteDZn_MD6cDmB1Q';

  // Nome da planilha
  const SHEET_NAME = 'Base';

  // 💙💙💙💙💙💙💙💙💙💙💙💙💙💙💙💙💙💙💙💙💙💙 SLIDE 3 💙💙💙💙💙💙💙💙💙💙💙💙💙💙💙💙💙💙💙💙💙💙

  const slideIndex = 2; // Índice do slide (baseado em 0)

  try {
    // Acessar a planilha e os dados
    var spreadsheet = SpreadsheetApp.openById(SHEETS_ID);
    var sheet = spreadsheet.getSheetByName(SHEET_NAME);
    

    // Acessar a apresentação e o slide
    var presentation = SlidesApp.openById(SLIDES_ID);
    var slide = presentation.getSlides()[slideIndex];

    // ############################## ⭐ PROJETOS NOVOS - 1ª FASE ⭐ ##############################################

        // Dados da Planilha - Colocar a célula da planilha
    var DATAMOUNTH = sheet.getRange('C2').getValue();
    var PN_INPROGRESS = sheet.getRange('B3').getValue();
    var PN_SUSPENDED = sheet.getRange('B4').getValue();
    var PN_TOTAL = sheet.getRange('B5').getValue();
    var PN_SUSPENDED_PERCENTAGE = sheet.getRange('B6').getValue();
    var PN_INPROGRESS_PERCENTAGE = sheet.getRange('B7').getValue();
    6
    slide.getShapes()[75].getText().setText(PN_INPROGRESS + "%"); // Em andamento
    slide.getShapes()[78].getText().setText(PN_SUSPENDED + "%"); // Suspensos
    slide.getShapes()[81].getText().setText(PN_TOTAL); // Total
    // slide.getShapes()[##].getText().setText(PN_SUSPENDED_PERCENTAGE  * 100 + "%"); // Suspensos
    // slide.getShapes()[##].getText().setText(PN_INPROGRESS_PERCENTAGE * 100 + "%"); // Suspensos
    slide.getShapes()[82].getText().setText(DATAMOUNTH); // Período


    // ##################################### ⭐ PROJETOS NOVOS - 2ª FASE ⭐ #######################################

  // Dados da Planilha - Colocar a célula da planilha
    var PN2_INPROGRESS = sheet.getRange('B15').getValue();
    var PN2_SUSPENDED = sheet.getRange('B16').getValue();
    var PN2_TOTAL = sheet.getRange('B17').getValue();
    var PN2_SUSPENDED_PERCENTAGE = sheet.getRange('B18').getValue();
    var PN2_INPROGRESS_PERCENTAGE = sheet.getRange('B19').getValue();

    slide.getShapes()[34].getText().setText(PN2_INPROGRESS + "%"); // Em andamento
    slide.getShapes()[37].getText().setText(PN2_SUSPENDED + "%"); // Suspensos
    slide.getShapes()[101].getText().setText(PN2_TOTAL); // Total
    // slide.getShapes()[##].getText().setText(PN2_SUSPENDED_PERCENTAGE + "%"); // Suspensos
    // slide.getShapes()[##].getText().setText(PN2_INPROGRESS_PERCENTAGE + "%"); // Suspensos


    // ##################################### ⭐ PROJETOS ANTIGOS - 1ª FASE e 2ª FASE ⭐ #######################################

   // Dados da Planilha - Colocar a célula da planilha
    var PA_INPROGRESS = sheet.getRange('B27').getValue();
    var PA_SUSPENDED = sheet.getRange('B28').getValue();
    var PA_TOTAL = sheet.getRange('B29').getValue();
    var PA_SUSPENDED_PERCENTAGE = sheet.getRange('B30').getValue();
    var PA_INPROGRESS_PERCENTAGE = sheet.getRange('B31').getValue();


    slide.getShapes()[40].getText().setText(PA_INPROGRESS + "%"); // Em andamento
    slide.getShapes()[43].getText().setText(PA_SUSPENDED + "%"); // Suspensos
    slide.getShapes()[104].getText().setText(PA_TOTAL); // Total
    // slide.getShapes()[##].getText().setText(PA_SUSPENDED_PERCENTAGE + "%"); // Suspensos
    // slide.getShapes()[##].getText().setText(PA_INPROGRESS_PERCENTAGE + "%"); // Suspensos



    // ##################################### ⭐ PROJETOS CONCLUÍDOS ⭐ #######################################

      var PC_FASE1 = sheet.getRange('B39').getValue();
      var PC_FASE2 = sheet.getRange('B40').getValue();
      var PC_TOTAL = sheet.getRange('B41').getValue();

      slide.getShapes()[22].getText().setText(PC_FASE1 + " - Fase 1"); // Projetos Concluídos - 1ª Fase 
      slide.getShapes()[83].getText().setText(PC_FASE2 + " -  Fase 2"); // Projetos Concluídos - 2ª Fase 
      slide.getShapes()[20].getText().setText(PC_TOTAL + " TOTAL"); // Projetos Concluídos - Total
    


    // ##################################### ⭐ PEP ⭐ #######################################

      var PEP_ATINGIDO = sheet.getRange('B49').getValue();
      var PEP_META = sheet.getRange('B50').getValue();
      

        slide.getShapes()[25].getText().setText(PEP_ATINGIDO + "%"); // PEP - Atingido
        // slide.getShapes()[26].getText().setText(PEP_META + "%"); // PEP - Meta

      // Ajustar a largura da barra de progresso com base na porcentagem
      var maxHeight_PEP = 100.55; // Largura máxima da barra de progresso
      var newHeight_PEP = Math.min((PEP_ATINGIDO / 100) * maxHeight_PEP + 0.000001, maxHeight_PEP)
      slide.getShapes()[24].setHeight(newHeight_PEP); // Ajustar a altura da barra de progresso



    // ##################################### ⭐ CONFORMIDADE DE PROJETOS ⭐ #######################################


      var CONFORMIDADE_ATINGIDO = sheet.getRange('B58').getValue();
      var CONFORMIDADE_META = sheet.getRange('B59').getValue();
      var CONFORMIDADE_COR_FONTE = sheet.getRange('L60').getValue();


        slide.getShapes()[71].getText().setText(CONFORMIDADE_ATINGIDO + "%"); // Conformidade - ATINGIDO
        // slide.getShapes()[72].getText().setText(CONFORMIDADE_META + "%"); // Conformidade - META
        slide.getShapes()[71].getText().getTextStyle().setForegroundColor(CONFORMIDADE_COR_FONTE); // Conformidade: Meta Atingida - Cor da fonte

      // Valor a ser puxado dentro da planilha
      var CRONOGRAMA = sheet.getRange('B60').getValue();
      var SINTONIA = sheet.getRange('B61').getValue();
      var APONTAMENTO_HORAS = sheet.getRange('B62').getValue();
      var ESCOPO = sheet.getRange('B63').getValue();
      var SIMULACAO = sheet.getRange('B64').getValue();
      var KICKOFF = sheet.getRange('B65').getValue();
      var ENCERRAMENTO = sheet.getRange('B66').getValue();


      // Cor de preenchimento da barra
      var BARRA_CRONOGRAMA = sheet.getRange('K60').getValue();
      var BARRA_SINTONIA = sheet.getRange('K61').getValue();
      var BARRA_APONTAMENTOH = sheet.getRange('K62').getValue();
      var BARRA_ESCOPO = sheet.getRange('K63').getValue();
      var BARRA_SIMULACAO = sheet.getRange('K64').getValue();
      var BARRA_KICKOFF = sheet.getRange('K65').getValue();
      var BARRA_ENCERRAMENTO = sheet.getRange('K66').getValue();
      
        // Ajustar os valores contidos dentro das barras
        slide.getShapes()[62].getText().setText(CRONOGRAMA + "%"); // Conformidade - Cronograma
        slide.getShapes()[60].getText().setText(SINTONIA + "%"); // Conformidade - Sintonia
        slide.getShapes()[58].getText().setText(APONTAMENTO_HORAS + "%"); // Conformidade - Apontamento de horas
        slide.getShapes()[68].getText().setText(ESCOPO + "%"); // Conformidade - Escopo
        slide.getShapes()[66].getText().setText(SIMULACAO + "%"); // Conformidade - Simulação
        slide.getShapes()[64].getText().setText(KICKOFF + "%"); // Conformidade - Kickoff
        slide.getShapes()[70].getText().setText(ENCERRAMENTO + "%"); // Conformidade - Encerramento

      
      // Ajustar a largura da barra de progresso com base na porcentagem
      var maxHeight_CRONOGRAMA = 68.26; // Largura máxima da barra de progresso (CRONOGRAMA)
      var newHeight_CRONOGRAMA = Math.min((CRONOGRAMA / 100) * maxHeight_CRONOGRAMA + 0.000001, maxHeight_CRONOGRAMA)
      slide.getShapes()[61].setHeight(newHeight_CRONOGRAMA); // Ajustar a altura da barra de progresso
      slide.getShapes()[61].getFill().setSolidFill(BARRA_CRONOGRAMA); // Aplicar a cor sólida ao preenchimento

      // Ajustar a largura da barra de progresso com base na porcentagem (SINTONIA)
      var maxHeight_SINTONIA = 68.26; // Largura máxima da barra de progresso
      var newHeight_SINTONIA = Math.min((SINTONIA / 100) * maxHeight_SINTONIA + 0.000001, maxHeight_SINTONIA)
      slide.getShapes()[59].setHeight(newHeight_SINTONIA); // Ajustar a altura da barra de progresso
      slide.getShapes()[59].getFill().setSolidFill(BARRA_SINTONIA); // Aplicar a cor sólida ao preenchimento

      // Ajustar a largura da barra de progresso com base na porcentagem (APONTAMENTO DE HORAS)
      var maxHeight_APONTAMENTO_HORAS = 68.26; // Largura máxima da barra de progresso
      var newHeight_APONTAMENTO_HORAS = Math.min((APONTAMENTO_HORAS / 100) * maxHeight_APONTAMENTO_HORAS + 0.000001, maxHeight_APONTAMENTO_HORAS)
      slide.getShapes()[57].setHeight(newHeight_APONTAMENTO_HORAS); // Ajustar a altura da barra de progresso
      slide.getShapes()[57].getFill().setSolidFill(BARRA_APONTAMENTOH); // Aplicar a cor sólida ao preenchimento

      // Ajustar a largura da barra de progresso com base na porcentagem (ESCOPO)
      var maxHeight_ESCOPO = 68.26; // Largura máxima da barra de progresso
      var newHeight_ESCOPO = Math.min((ESCOPO / 100) * maxHeight_ESCOPO + 0.000001, maxHeight_ESCOPO)
      slide.getShapes()[67].setHeight(newHeight_ESCOPO); // Ajustar a altura da barra de progresso
      slide.getShapes()[67].getFill().setSolidFill(BARRA_ESCOPO); // Aplicar a cor sólida ao preenchimento

      // Ajustar a largura da barra de progresso com base na porcentagem (SIMULAÇÃO)
      var maxHeight_SIMULACAO = 68.26; // Largura máxima da barra de progresso
      var newHeight_SIMULACAO = Math.min((SIMULACAO / 100) * maxHeight_SIMULACAO + 0.000001, maxHeight_SIMULACAO)
      slide.getShapes()[65].setHeight(newHeight_SIMULACAO); // Ajustar a altura da barra de progresso
      slide.getShapes()[65].getFill().setSolidFill(BARRA_SIMULACAO); // Aplicar a cor sólida ao preenchimento

      // Ajustar a largura da barra de progresso com base na porcentagem (KICKOFF)
      var maxHeight_KICKOFF = 68.26; // Largura máxima da barra de progresso
      var newHeight_KICKOFF = Math.min((KICKOFF / 100) * maxHeight_KICKOFF + 0.000001, maxHeight_KICKOFF)
      slide.getShapes()[63].setHeight(newHeight_KICKOFF); // Ajustar a altura da barra de progresso
      slide.getShapes()[63].getFill().setSolidFill(BARRA_KICKOFF); // Aplicar a cor sólida ao preenchimento

      // Ajustar a largura da barra de progresso com base na porcentagem (ENCERRAMENTO)
      var maxHeight_ENCERRAMENTO = 68.26; // Largura máxima da barra de progresso
      var newHeight_ENCERRAMENTO = Math.min((ENCERRAMENTO / 100) * maxHeight_ENCERRAMENTO + 0.000001, maxHeight_ENCERRAMENTO)
      slide.getShapes()[69].setHeight(newHeight_ENCERRAMENTO); // Ajustar a altura da barra de progresso
      slide.getShapes()[69].getFill().setSolidFill(BARRA_ENCERRAMENTO); // Aplicar a cor sólida ao preenchimento
      

    // ##################################### ⭐ PROJETOS CONCLUÍDOS - CLUSTER ⭐ #######################################


      var CLUSTER_500_A = sheet.getRange('B74').getValue();
      var CLUSTER_1000_A = sheet.getRange('B76').getValue();
      var CLUSTER_1500_A = sheet.getRange('B78').getValue();
      var CLUSTER_1500MAIS_A = sheet.getRange('B80').getValue();
      var CLUSTER_TOTAL_A = sheet.getRange('B83').getValue();
      
      slide.getShapes()[31].getText().setText(CLUSTER_500_A); // Cluster <= 500H (ANTIGOS)
      slide.getShapes()[91].getText().setText(CLUSTER_1000_A); // Cluster <= 1000H (ANTIGOS)
      slide.getShapes()[94].getText().setText(CLUSTER_1500_A); // Cluster <= 1500H (ANTIGOS)
      slide.getShapes()[97].getText().setText(CLUSTER_1500MAIS_A); // Cluster > 1500H (ANTIGOS)
      slide.getShapes()[14].getText().setText(CLUSTER_TOTAL_A); // Cluster TOTAL (ANTIGOS)

      var CLUSTER_500_N = sheet.getRange('B75').getValue();
      var CLUSTER_1000_N = sheet.getRange('B77').getValue();
      var CLUSTER_1500_N = sheet.getRange('B79').getValue();
      var CLUSTER_1500MAIS_N = sheet.getRange('B81').getValue();
      var CLUSTER_TOTAL_N = sheet.getRange('B83').getValue();

      slide.getShapes()[44].getText().setText(CLUSTER_500_N); // Cluster <= 500H (NOVOS)
      slide.getShapes()[92].getText().setText(CLUSTER_1000_N); // Cluster <= 1000H (NOVOS)
      slide.getShapes()[95].getText().setText(CLUSTER_1500_N); // Cluster <= 1500H (NOVOS)
      slide.getShapes()[98].getText().setText(CLUSTER_1500MAIS_N); // Cluster > 1500H (NOVOS)
      slide.getShapes()[19].getText().setText(CLUSTER_TOTAL_N); // Cluster TOTAL (NOVOS)


    // ##################################### ⭐ DMP ⭐ #######################################





    // ❎❎❎❎❎❎❎❎❎❎❎❎❎❎❎❎❎❎❎❎ SLIDE 4 ❎❎❎❎❎❎❎❎❎❎❎❎❎❎❎❎❎❎❎❎

      const slideIndex2 = 3; // Índice do slide (baseado em 0)

    // Acessar a apresentação e o slide
    var presentation = SlidesApp.openById(SLIDES_ID);
    var slide2 = presentation.getSlides()[slideIndex2];

        var DATAMOUNTH2 = sheet.getRange('C2').getValue();
      slide2.getShapes()[24].getText().setText(DATAMOUNTH2); // Período


      var SATISFAÇAO_ATINGIDO_PN = sheet.getRange('B106').getValue();
      var SATISFAÇAO_META_PN = sheet.getRange('B107').getValue();
      var SATISFAÇAO_COR_FONTE_PN = sheet.getRange('L106').getValue();

      var SATISFAÇAO_ATINGIDO_PA = sheet.getRange('B108').getValue();
      var SATISFAÇAO_META_PA = sheet.getRange('B109').getValue();
      var SATISFAÇAO_COR_FONTE_PA = sheet.getRange('L107').getValue();

      // Cor de preenchimento da barra
      var BARRA_SAT_NOVOS = sheet.getRange('K106').getValue();
      var BARRA_SAT_ANTIGOS = sheet.getRange('K107').getValue();


        // Ajustar os valores  - SATISFAÇÃO DE PROJETOS NOVOS
        slide2.getShapes()[5].getText().setText(SATISFAÇAO_ATINGIDO_PN + "%"); // SATISFAÇÃO - ATINGIDO
        // slide.getShapes()[6].getText().setText(SATISFAÇAO_META_PN + "%"); // SATISFAÇÃO - META
        slide2.getShapes()[5].getText().getTextStyle().setForegroundColor(SATISFAÇAO_COR_FONTE_PN); // Satisfação: Atingido - Cor da fonte

        // Ajustar os valores  - SATISFAÇÃO DE PROJETOS ANTIGOS
        slide2.getShapes()[9].getText().setText(SATISFAÇAO_ATINGIDO_PA + "%"); // SATISFAÇÃO - ATINGIDO
        // slide.getShapes()[10].getText().setText(SATISFAÇAO_META_PA + "%"); // SATISFAÇÃO - META
        slide2.getShapes()[9].getText().getTextStyle().setForegroundColor(SATISFAÇAO_COR_FONTE_PA); // Satisfação: Meta - Cor da fonte


        // Ajustar a largura da barra de progresso com base na porcentagem (SATISFAÇÃO - PROJETOS NOVOS)
      var maxHeight_SAT_PN = 100.65; // Largura máxima da barra de progresso
      var newHeight_SAT_PN = Math.min((SATISFAÇAO_ATINGIDO_PN / 100) * maxHeight_SAT_PN + 0.000001, maxHeight_SAT_PN)
      slide2.getShapes()[4].setHeight(newHeight_SAT_PN); // Ajustar a altura da barra de progresso
      slide2.getShapes()[4].getFill().setSolidFill(BARRA_SAT_NOVOS); // Aplicar a cor sólida ao preenchimento

        // Ajustar a largura da barra de progresso com base na porcentagem (SATISFAÇÃO - PROJETOS ANTIGOS)
      var maxHeight_SAT_PA = 100.65; // Largura máxima da barra de progresso
      var newHeight_SAT_PA = Math.min((SATISFAÇAO_ATINGIDO_PA / 100) * maxHeight_SAT_PA + 0.000001, maxHeight_SAT_PA)
      slide2.getShapes()[8].setHeight(newHeight_SAT_PA); // Ajustar a altura da barra de progresso
      slide2.getShapes()[8].getFill().setSolidFill(BARRA_SAT_ANTIGOS); // Aplicar a cor sólida ao preenchimento

    // Ofensores - Lista

      var OFENSOR_01 = sheet.getRange('B117').getValue();
      var OFENSOR_02 = sheet.getRange('B118').getValue();
      var OFENSOR_03 = sheet.getRange('B119').getValue();
      var OFENSOR_04 = sheet.getRange('B120').getValue();
      var OFENSOR_05 = sheet.getRange('B121').getValue();

      slide2.getShapes()[13].getText().setText(OFENSOR_01); // Item 1 da lista de Ofensores
      slide2.getShapes()[14].getText().setText(OFENSOR_02); // Item 2 da lista de Ofensores
      slide2.getShapes()[16].getText().setText(OFENSOR_03); // Item 3 da lista de Ofensores
      slide2.getShapes()[18].getText().setText(OFENSOR_04); // Item 4 da lista de Ofensores
      slide2.getShapes()[20].getText().setText(OFENSOR_05); // Item 5 da lista de Ofensores



      // ⚡⚡⚡⚡⚡⚡⚡⚡⚡⚡⚡⚡⚡⚡⚡⚡⚡⚡⚡ SLIDE 5 ⚡⚡⚡⚡⚡⚡⚡⚡⚡⚡⚡⚡⚡⚡⚡⚡⚡⚡⚡


    // Indice do slide 5
    const slideIndex3 = 4; // Índice do slide (baseado em 0)

    // Acessar a apresentação e o slide
    var presentation = SlidesApp.openById(SLIDES_ID);
    var slide3 = presentation.getSlides()[slideIndex3];

    var DATAMOUNTH2 = sheet.getRange('C2').getValue();
    slide3.getShapes()[37].getText().setText(DATAMOUNTH2); // Período

    // Plano de ação - caixas de texto
      var PLAN_AC_EA = sheet.getRange('B138').getValue();
      var PLAN_AC_ENC = sheet.getRange('B139').getValue();

      slide3.getShapes()[33].getText().setText(PLAN_AC_EA); // Plano de ação - Em Aberto
      slide3.getShapes()[36].getText().setText(PLAN_AC_ENC); // Plano de ação - Encerrados

    // Reclamações - caixa de texto
      var RECL_AC_EA = sheet.getRange('B148').getValue();
      var TOTAL_PORTF = sheet.getRange('B149').getValue();

      slide3.getShapes()[5].getText().setText(RECL_AC_EA); // // Reclamações - Em Aberto
      slide3.getShapes()[8].getText().setText(RECL_AC_EA); // // Reclamações - Em Aberto
      slide3.getShapes()[3].getText().setText(TOTAL_PORTF); // // Reclamações - Total do Portfólio

      // MMR - caixa de texto
      var VLR_MMR = sheet.getRange('B158').getValue();
      var VLR_MMR_Formatado = 'R$ ' + VLR_MMR.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

      slide3.getShapes()[18].getText().setText(VLR_MMR_Formatado); // MMR - Total
      

      // Principais clientes
      var CL1 = sheet.getRange('B167').getValue();
      var CL2 = sheet.getRange('B169').getValue();
      var CL3 = sheet.getRange('B171').getValue();
      var CL4 = sheet.getRange('B173').getValue();
      var CL5 = sheet.getRange('B175').getValue();

      slide3.getShapes()[9].getText().setText(CL1); // Clientes / Unidades (1)
      slide3.getShapes()[10].getText().setText(CL2); // Clientes / Unidades (2)
      slide3.getShapes()[12].getText().setText(CL3); // Clientes / Unidades (3)
      slide3.getShapes()[14].getText().setText(CL4); // Clientes / Unidades (4)
      slide3.getShapes()[16].getText().setText(CL5); // Clientes / Unidades (5)

      // Principais clientes - MMR em Risco (R$)

      var CL1_MMR = sheet.getRange('B168').getValue();
      var CL2_MMR = sheet.getRange('B170').getValue();
      var CL3_MMR = sheet.getRange('B172').getValue();
      var CL4_MMR = sheet.getRange('B174').getValue();
      var CL5_MMR = sheet.getRange('B176').getValue();

      slide3.getShapes()[23].getText().setText(CL1_MMR + " mil"); // MMR por Cliente / Unidade (1)
      slide3.getShapes()[24].getText().setText(CL2_MMR + " mil"); // MMR por Cliente / Unidade (2)
      slide3.getShapes()[25].getText().setText(CL3_MMR + " mil"); // MMR por Cliente / Unidade (3)
      slide3.getShapes()[26].getText().setText(CL4_MMR + " mil"); // MMR por Cliente / Unidade (4)
      slide3.getShapes()[27].getText().setText(CL5_MMR + " mil"); // MMR por Cliente / Unidade (5)

      // 💚💚💚💚💚💚💚💚💚💚💚💚💚💚💚💚💚💚💚 SLIDE 6 💜💜💜💜💜💜💜💜💜💜💜💜💜💜💜💜💜💜💜💜


        const slideIndex4 = 5; // Índice do slide (baseado em 0)

    // Acessar a apresentação e o slide
    var presentation = SlidesApp.openById(SLIDES_ID);
    var slide4 = presentation.getSlides()[slideIndex4];

    // Dados da planilha para o slide
    var DATAMOUNTH3 = sheet.getRange('C2').getValue();
    slide4.getShapes()[12].getText().setText(DATAMOUNTH3); // Período


  } catch (e) {
    Logger.log(e.toString());
  }
}

// Botão "Atualizar" no Slide
function onOpen() {
  var ui = SlidesApp.getUi();
  ui.createMenu('Atualizar')
    .addItem('Atualizar base', 'updateSlidesFromSheet')
    .addToUi();
}