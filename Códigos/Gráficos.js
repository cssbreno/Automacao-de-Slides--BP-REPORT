// Gráfico 1

function createCustomDonutChart() {
  const sheet = SpreadsheetApp.openById(SHEETS_ID).getSheetByName(SHEET_NAME);
  const range = sheet.getRange("B6:B7"); // Ajuste o intervalo conforme seus dados

  // Criação do gráfico de rosca
  const chart = sheet
    .newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(range)
    .setPosition(5, 5, 0, 0)
    .setOption("pieHole", 0.4)
    .setOption("pieSliceText", "value") // Exibe os valores dentro das fatias
    .setOption("legend", { position: "none" }) // Remove a legenda externa
    .setOption("backgroundColor", { fill: "#FFFFFF" }) // Define o fundo como transparente
    .setOption("pieSliceTextStyle", {
      fontSize: "24",
      color: "#FFFFFF",
      bold: true,
    }) // Estiliza o texto das fatias
    .setOption("slices", {
      0: { offset: 0.1, color: "#FF0000" }, // Configura a cor e o deslocamento para a primeira fatia
      1: { offset: 0.1, color: "#274D8D" }, // Configura a cor e o deslocamento para a segunda fatia
    })
    .setOption("pieSliceBorderColor", "#D3D3D3") // Define a cor da borda das fatias (cinza claro)
    .setOption("pieSliceBorderWidth", 2) // Define a largura da borda
    .setOption("pieSliceBorderRadius", 5) // Define o raio da borda (para efeito de arredondamento, se aplicável)
    .build();

  sheet.insertChart(chart);
}

function insertChartIntoSlide(chartIndex) {
  const slides = SlidesApp.openById(SLIDES_ID);
  const slide = slides.getSlides()[slideIndex];
  const sheet = SpreadsheetApp.openById(SHEETS_ID);
  const charts = sheet.getSheets()[0].getCharts();

  if (charts.length > 0) {
    const chart = charts[0];
    const blob = chart.getAs("image/png");
    slide.insertImage(blob);
  }
}

function updateChart(imageIndex) {
  const slides = SlidesApp.openById(SLIDES_ID);
  const slide = slides.getSlides()[slideIndex];
  const sheet = SpreadsheetApp.openById(SHEETS_ID);
  const charts = sheet.getSheets()[0].getCharts();

  if (charts.length > imageIndex) {
    const chart = charts[imageIndex];
    const blob = chart.getAs("image/png");

    const images = slide.getImages();

    if (images.length > imageIndex) {
      const image = images[imageIndex];

      const left = image.getLeft();
      const top = image.getTop();
      const width = image.getWidth();
      const height = image.getHeight();

      image.remove();
      slide.insertImage(blob, left, top, width, height);
    } else {
      slide.insertImage(blob);
    }
  }
}

// Gráfico 2

function createCustomDonutChart2() {
  const sheet = SpreadsheetApp.openById(SHEETS_ID).getSheetByName(SHEET_NAME);
  const range = sheet.getRange("B18:B19"); // Ajuste o intervalo conforme seus dados

  // Criação do gráfico de rosca
  const chart = sheet
    .newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(range)
    .setPosition(5, 5, 0, 0)
    .setOption("pieHole", 0.4)
    .setOption("pieSliceText", "value") // Exibe os valores dentro das fatias
    .setOption("legend", { position: "none" }) // Remove a legenda externa
    .setOption("backgroundColor", { fill: "#FFFFFF" }) // Define o fundo como transparente
    .setOption("pieSliceTextStyle", {
      fontSize: "24",
      color: "#FFFFFF",
      bold: true,
    }) // Estiliza o texto das fatias
    .setOption("slices", {
      0: { offset: 0.1, color: "#FF0000" }, // Configura a cor e o deslocamento para a primeira fatia
      1: { offset: 0.1, color: "#274D8D" }, // Configura a cor e o deslocamento para a segunda fatia
    })
    .setOption("pieSliceBorderColor", "#D3D3D3") // Define a cor da borda das fatias (cinza claro)
    .setOption("pieSliceBorderWidth", 2) // Define a largura da borda
    .setOption("pieSliceBorderRadius", 5) // Define o raio da borda (para efeito de arredondamento, se aplicável)
    .build();

  sheet.insertChart(chart);
}

function insertChartIntoSlide2() {
  insertChartIntoSlide(1); // Inserir o segundo gráfico
}

function updateChart2() {
  updateChart(1); // Atualizar a imagem correspondente ao segundo gráfico
}

// Gráfico 3

function createCustomDonutChart3() {
  const sheet = SpreadsheetApp.openById(SHEETS_ID).getSheetByName(SHEET_NAME);
  const range = sheet.getRange("B30:B31"); // Ajuste o intervalo conforme seus dados

  // Criação do gráfico de rosca
  const chart = sheet
    .newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(range)
    .setPosition(5, 5, 0, 0)
    .setOption("pieHole", 0.4)
    .setOption("pieSliceText", "value") // Exibe os valores dentro das fatias
    .setOption("legend", { position: "none" }) // Remove a legenda externa
    .setOption("backgroundColor", { fill: "#FFFFFF" }) // Define o fundo como transparente
    .setOption("pieSliceTextStyle", {
      fontSize: "24",
      color: "#FFFFFF",
      bold: true,
    }) // Estiliza o texto das fatias
    .setOption("slices", {
      0: { offset: 0.1, color: "#FF0000" }, // Configura a cor e o deslocamento para a primeira fatia
      1: { offset: 0.1, color: "#274D8D" }, // Configura a cor e o deslocamento para a segunda fatia
    })
    .setOption("pieSliceBorderColor", "#D3D3D3") // Define a cor da borda das fatias (cinza claro)
    .setOption("pieSliceBorderWidth", 2) // Define a largura da borda
    .setOption("pieSliceBorderRadius", 5) // Define o raio da borda (para efeito de arredondamento, se aplicável)
    .build();

  sheet.insertChart(chart);
}

function insertChartIntoSlide3() {
  insertChartIntoSlide(2); // Inserir o terceiro gráfico
}

function updateChart3() {
  updateChart(2); // Atualizar a imagem correspondente ao terceiro gráfico
}

// Menu

function onOpen() {
  const ui = SlidesApp.getUi();
  ui.createMenu("Planplan")
    .addItem("Atualizar Gráfico 1", "updateChart")
    .addItem("Atualizar Gráfico 2", "updateChart2")
    .addItem("Atualizar Gráfico 3", "updateChart3")
    .addToUi();
}
