  // Nome da planilha
  const SHEET_RO = 'BP RONDÔNIA';
  const slideIndex8 = 7; // Índice do slide (baseado em 0)

// Gráfico 1
function createCustomDonutChart() {
  const sheet = SpreadsheetApp.openById(SHEETS_ID).getSheetByName(SHEET_RO);
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
  insertChartIntoSlide(chart);
}

// Gráfico 2
function createCustomDonutChart2() {
  const sheet = SpreadsheetApp.openById(SHEETS_ID).getSheetByName(SHEET_RO);
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
  insertChartIntoSlide(chart);
}

// Gráfico 3
function createCustomDonutChart3() {
  const sheet = SpreadsheetApp.openById(SHEETS_ID).getSheetByName(SHEET_RO);
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
  insertChartIntoSlide(chart);
}

// Função para inserir gráficos no slide
function insertChartIntoSlide(chart) {
  if (chart) { 
    const slides = SlidesApp.openById(SLIDES_ID);
    const slide = slides.getSlides()[slideIndex8];
    const blob = chart.getAs('image/png'); 
    slide.insertImage(blob);
  }
}

// Função para atualizar os três últimos gráficos
function updateLastThreeCharts() {
  const slides = SlidesApp.openById(SLIDES_ID);
  const slide = slides.getSlides()[slideIndex8];
  const images = slide.getImages();

  if (images.length >= 3) {
    const image1 = images[images.length - 3];
    const image2 = images[images.length - 2];
    const image3 = images[images.length - 1];
    
    const sheet = SpreadsheetApp.openById(SHEETS_ID).getSheetByName(SHEET_RO);
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

// Menu
function onOpen() {
  const ui = SlidesApp.getUi();
  ui.createMenu('Planplan')
    .addItem('Atualizar Gráficos', 'updateLastThreeCharts')
    .addToUi();
}
