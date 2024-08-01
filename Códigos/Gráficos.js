function createCustomDonutChart() {
    const sheet = SpreadsheetApp.openById(SHEETS_ID).getSheetByName(SHEET_NAME);
    const range = sheet.getRange('B6:B7'); // Ajuste o intervalo conforme seus dados
    
    // Criação do gráfico de rosca
    const chart = sheet.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(range)
      .setPosition(5, 5, 0, 0)
      .setOption('pieHole', 0.4)
      .setOption('pieSliceText', 'value') // Exibe os valores dentro das fatias
      .setOption('legend', { position: 'none' }) // Remove a legenda externa
      .setOption('backgroundColor', { 'fill': '#FFFFFF' }) // Define o fundo como transparente
      .setOption('pieSliceTextStyle', { 'fontSize': '24','color': '#FFFFFF', 'bold': true }) // Estiliza o texto das fatias
      .setOption('slices', {
        0: { offset: 0.1, color: '#FF0000' }, // Configura a cor e o deslocamento para a primeira fatia
        1: { offset: 0.1, color: '#2B4A76' }  // Configura a cor e o deslocamento para a segunda fatia
      })
      .setOption('pieSliceBorderColor', '#D3D3D3') // Define a cor da borda das fatias (cinza claro)
      .setOption('pieSliceBorderWidth', 2) // Define a largura da borda
      .setOption('pieSliceBorderRadius', 5) // Define o raio da borda (para efeito de arredondamento, se aplicável)
      .build();
    
    sheet.insertChart(chart);
  }
  
  
  function insertChartIntoSlide() {
    const slides = SlidesApp.openById(SLIDES_ID);
    const slide = slides.getSlides()[slideIndex];
    const sheet = SpreadsheetApp.openById(SHEETS_ID);
    const charts = sheet.getSheets()[0].getCharts();
    
    if (charts.length > 0) {
      const chart = charts[0];
      const blob = chart.getAs('image/png');
      slide.insertImage(blob);
    }
  }
  
  
  function updateChart() {
    const slides = SlidesApp.openById(SLIDES_ID);
    const slide = slides.getSlides()[slideIndex];
    const sheet = SpreadsheetApp.openById(SHEETS_ID);
    const charts = sheet.getSheets()[0].getCharts();
    
    if (charts.length > 0) {
      const chart = charts[0];
      const blob = chart.getAs('image/png');
      
      // Remover imagens antigas, se houver
      const images = slide.getImages();
      for (const image of images) {
        image.remove();
      }
      
      // Adicionar a nova imagem
      slide.insertImage(blob);
    }
  }
  
  
  function onOpen() {
    const ui = SlidesApp.getUi();
    ui.createMenu('Planplan')
      .addItem('Atualizar Gráfico', 'updateChart')
      .addToUi();
  }
  