    // IDs da apresentação e da planilha
    const SLIDES_ID = '1whKu0ncTiW3qFmPUXusqTL7zu_ai_uy4kz09oyoaq34';
    const SHEETS_ID = '1GZAPqryb7C3J_rF2GpyMAYkqj5itteDZn_MD6cDmB1Q';
  
    // Nome da planilha
    const SHEET_BP_SALVADOR = 'BP SALVADOR';
    const slideIndex9 = 8; // Índice do slide (baseado em 0)
  
  // Gráfico 1
  function createChartBPSALVADOR1() {
    const sheet = SpreadsheetApp.openById(SHEETS_ID).getSheetByName(SHEET_BP_SALVADOR);
    const range = sheet.getRange('B6:B7'); // Ajuste o intervalo conforme seus dados
  
    const chart = sheet.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(range)
      .setPosition(5, 5, 0, 0)
      .setOption('pieHole', 0.4)
      .setOption('pieSliceText', 'value')
      .setOption('legend', { position: 'none' })
      .setOption('backgroundColor', { 'fill': '#FFFFFF' })
      .setOption('pieSliceTextStyle', { 'fontSize': '24','color': '#FFFFFF', 'bold': true })
      .setOption('slices', {
        0: { offset: 0.1, color: '#FF0000' },
        1: { offset: 0.1, color: '#274D8D' }
      })
      .setOption('pieSliceBorderColor', '#D3D3D3')
      .setOption('pieSliceBorderWidth', 2)
      .setOption('pieSliceBorderRadius', 5)
      .build();
  
    sheet.insertChart(chart);
    insertChartBPSALVADOR(chart);
  }
  
  // Gráfico 2
  function createChartBPSALVADOR2() {
    const sheet = SpreadsheetApp.openById(SHEETS_ID).getSheetByName(SHEET_BP_SALVADOR);
    const range = sheet.getRange('B18:B19'); // Ajuste o intervalo conforme seus dados
  
    const chart = sheet.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(range)
      .setPosition(5, 5, 0, 0)
      .setOption('pieHole', 0.4)
      .setOption('pieSliceText', 'value')
      .setOption('legend', { position: 'none' })
      .setOption('backgroundColor', { 'fill': '#FFFFFF' })
      .setOption('pieSliceTextStyle', { 'fontSize': '24','color': '#FFFFFF', 'bold': true })
      .setOption('slices', {
        0: { offset: 0.1, color: '#FF0000' },
        1: { offset: 0.1, color: '#274D8D' }
      })
      .setOption('pieSliceBorderColor', '#D3D3D3')
      .setOption('pieSliceBorderWidth', 2)
      .setOption('pieSliceBorderRadius', 5)
      .build();
  
    sheet.insertChart(chart);
    insertChartBPSALVADOR(chart);
  }
  
  // Gráfico 3
  function createChartBPSALVADOR3() {
    const sheet = SpreadsheetApp.openById(SHEETS_ID).getSheetByName(SHEET_BP_SALVADOR);
    const range = sheet.getRange('B30:B31'); // Ajuste o intervalo conforme seus dados
  
    const chart = sheet.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(range)
      .setPosition(5, 5, 0, 0)
      .setOption('pieHole', 0.4)
      .setOption('pieSliceText', 'value')
      .setOption('legend', { position: 'none' })
      .setOption('backgroundColor', { 'fill': '#FFFFFF' })
      .setOption('pieSliceTextStyle', { 'fontSize': '24','color': '#FFFFFF', 'bold': true })
      .setOption('slices', {
        0: { offset: 0.1, color: '#FF0000' },
        1: { offset: 0.1, color: '#274D8D' }
      })
      .setOption('pieSliceBorderColor', '#D3D3D3')
      .setOption('pieSliceBorderWidth', 2)
      .setOption('pieSliceBorderRadius', 5)
      .build();
  
    sheet.insertChart(chart);
    insertChartBPSALVADOR(chart);
  }
  
  // Função para inserir gráficos no slide
  function insertChartBPSALVADOR(chart) {
    if (chart) { 
      const slides = SlidesApp.openById(SLIDES_ID);
      const slide = slides.getSlides()[slideIndex9];
      const blob = chart.getAs('image/png'); 
      slide.insertImage(blob);
    }
  }
  
  // Função para atualizar os três últimos gráficos
  function updateChartsSALV() {
    const slides = SlidesApp.openById(SLIDES_ID);
    const slide = slides.getSlides()[slideIndex9];
    const images = slide.getImages();
  
    if (images.length >= 3) {
      const image1 = images[images.length - 3];
      const image2 = images[images.length - 2];
      const image3 = images[images.length - 1];
      
      const sheet = SpreadsheetApp.openById(SHEETS_ID).getSheetByName(SHEET_BP_SALVADOR);
      const charts = sheet.getCharts();
  
      if (charts.length >= 3) {
        updateImage(slide, image1, charts[charts.length - 3]);
        updateImage(slide, image2, charts[charts.length - 2]);
        updateImage(slide, image3, charts[charts.length - 1]);
      }
    }
  }
  
  // Função para atualizar uma imagem no slide com o gráfico correspondente
  function updateImage(slide, image, chart) {
    const blob = chart.getAs('image/png');
    const left = image.getLeft();
    const top = image.getTop();
    const width = image.getWidth();
    const height = image.getHeight();
  
    image.remove();
    slide.insertImage(blob, left, top, width, height);
  }