// Gráfico 1

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
      1: { offset: 0.1, color: '#274D8D' }  // Configura a cor e o deslocamento para a segunda fatia
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
    
    // Encontrar e substituir a penúltima imagem
    const images = slide.getImages();
    
    if (images.length > 1) {
      const imageIndex = images.length - 3; // Índice da penúltima imagem
      
      // Seleciona a penúltima imagem adicionada
      const image = images[imageIndex];
      
      // Armazena a posição e o tamanho da imagem
      const left = image.getLeft();
      const top = image.getTop();
      const width = image.getWidth();
      const height = image.getHeight();
      
      // Remove a penúltima imagem
      image.remove();
      
      // Adiciona a nova imagem na mesma posição e tamanho
      slide.insertImage(blob, left, top, width, height);
    } else if (images.length > 0) {
      // Caso haja menos de 2 imagens, apenas adiciona a nova imagem
      slide.insertImage(blob);
    }
  }
}



// Gráfico 2

function createCustomDonutChart2() {
  const sheet = SpreadsheetApp.openById(SHEETS_ID).getSheetByName(SHEET_NAME);
  const range = sheet.getRange('B18:B19'); // Ajuste o intervalo conforme seus dados
  
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
      1: { offset: 0.1, color: '#274D8D' }  // Configura a cor e o deslocamento para a segunda fatia
    })
    .setOption('pieSliceBorderColor', '#D3D3D3') // Define a cor da borda das fatias (cinza claro)
    .setOption('pieSliceBorderWidth', 2) // Define a largura da borda
    .setOption('pieSliceBorderRadius', 5) // Define o raio da borda (para efeito de arredondamento, se aplicável)
    .build();
  
  sheet.insertChart(chart);
}


function insertChartIntoSlide2() {
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

function updateChart2() {
  const slides = SlidesApp.openById(SLIDES_ID);
  const slide = slides.getSlides()[slideIndex];
  const sheet = SpreadsheetApp.openById(SHEETS_ID);
  const charts = sheet.getSheets()[0].getCharts();
  
  if (charts.length > 0) {
    const chart = charts[0];
    const blob = chart.getAs('image/png');
    
    // Encontrar e substituir a última imagem
    const images = slide.getImages();
    
    if (images.length > 2) {
      // Seleciona a última imagem adicionada
      const image = images[images.length - 2];
      
      // Armazena a posição e o tamanho da imagem
      const left = image.getLeft();
      const top = image.getTop();
      const width = image.getWidth();
      const height = image.getHeight();
      
      // Remove a última imagem
      image.remove();
      
      // Adiciona a nova imagem na mesma posição e tamanho
      slide.insertImage(blob, left, top, width, height);
    } else {
      // Caso não haja imagens no slide, apenas adiciona a nova imagem
      slide.insertImage(blob);
    }
  }
}

// Gráfico 3

function createCustomDonutChart3() {
  const sheet = SpreadsheetApp.openById(SHEETS_ID).getSheetByName(SHEET_NAME);
  const range = sheet.getRange('B30:B31'); // Ajuste o intervalo conforme seus dados
  
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
      1: { offset: 0.1, color: '#274D8D' }  // Configura a cor e o deslocamento para a segunda fatia
    })
    .setOption('pieSliceBorderColor', '#D3D3D3') // Define a cor da borda das fatias (cinza claro)
    .setOption('pieSliceBorderWidth', 2) // Define a largura da borda
    .setOption('pieSliceBorderRadius', 5) // Define o raio da borda (para efeito de arredondamento, se aplicável)
    .build();
  
  sheet.insertChart(chart);
}


function insertChartIntoSlide3() {
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

function updateChart3() {
  const slides = SlidesApp.openById(SLIDES_ID);
  const slide = slides.getSlides()[slideIndex];
  const sheet = SpreadsheetApp.openById(SHEETS_ID);
  const charts = sheet.getSheets()[0].getCharts();
  
  if (charts.length > 0) {
    const chart = charts[0];
    const blob = chart.getAs('image/png');
    
    // Encontrar e substituir a última imagem
    const images = slide.getImages();
    
    if (images.length > 3) {
      // Seleciona a última imagem adicionada
      const image = images[images.length - 3];
      
      // Armazena a posição e o tamanho da imagem
      const left = image.getLeft();
      const top = image.getTop();
      const width = image.getWidth();
      const height = image.getHeight();
      
      // Remove a última imagem
      image.remove();
      
      // Adiciona a nova imagem na mesma posição e tamanho
      slide.insertImage(blob, left, top, width, height);
    } else {
      // Caso não haja imagens no slide, apenas adiciona a nova imagem
      slide.insertImage(blob);
    }
  }
}


function onOpen() {
  const ui = SlidesApp.getUi();
  ui.createMenu('Planplan')
    .addItem('Atualizar Gráfico', 'updateChart')
    .addToUi();
}
