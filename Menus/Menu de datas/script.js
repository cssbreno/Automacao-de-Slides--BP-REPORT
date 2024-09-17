// Função que cria o menu e exibe o seletor de data
function onOpen() {
  const ui = SlidesApp.getUi();
  ui.createMenu("Período de Referência")
    .addItem("Selecionar Mês/Ano", "showDateSelector") // Associa a função 'showDateSelector' ao item de menu
    .addToUi();
}

// Função que exibe o seletor de data
function showDateSelector() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile("DateSelector")
    .setWidth(400)
    .setHeight(350);
  SlidesApp.getUi().showModalDialog(htmlOutput, "Selecionar data");
}

// Função que atualiza os slides com a data formatada
function updateSlideWithDate(dateStr) {
  if (!validateDate(dateStr)) {
    SlidesApp.getUi().alert("Formato inválido. Use o formato MM/AA.");
    return;
  }

  const formattedDate = convertDateToMonthYear(dateStr);
  const SLIDES_ID = "1whKu0ncTiW3qFmPUXusqTL7zu_ai_uy4kz09oyoaq34";

  const slideIndices = [
    6, 7, 12, 17, 22, 27, 32, 37, 38, 43, 48, 53, 58, 63, 68, 69, 74, 79, 84,
    89, 94, 99, 100, 105, 110, 115, 120, 125, 130, 135, 136, 141, 146, 151, 156,
  ];

  try {
    var presentation = SlidesApp.openById(SLIDES_ID);
    var slides = presentation.getSlides();

    slideIndices.forEach(function (slideIndex) {
      var slide = slides[slideIndex];
      var shapes = slide.getShapes();
      var shapeIndex = 1; // Supondo que o shape seja o segundo na ordem (baseado em 0)

      if (shapes[shapeIndex]) {
        var textRange = shapes[shapeIndex].getText();
        textRange.clear(); // Limpa o conteúdo anterior
        textRange.setText(formattedDate); // Define o novo texto no formato Mês/AA
      } else {
        SlidesApp.getUi().alert(
          "A caixa de texto não foi encontrada no slide de índice: " +
            slideIndex
        );
      }
    });
  } catch (e) {
    Logger.log("Erro ao atualizar os slides: " + e.message);
  }
}

// Função que converte a data de MM/AA para Mês/AA
function convertDateToMonthYear(dateStr) {
  const [month, year] = dateStr.split("/");

  const months = [
    "Janeiro",
    "Fevereiro",
    "Março",
    "Abril",
    "Maio",
    "Junho",
    "Junho",
    "Agosto",
    "Setembro",
    "Outubro",
    "Novembro",
    "Dezembro",
  ];

  const monthName = months[parseInt(month) - 1];
  return `${monthName}/${year}`;
}

// Função para validar o formato MM/AA
function validateDate(dateStr) {
  const regExp = /^(0[1-9]|1[0-2])\/\d{2}$/;
  return regExp.test(dateStr);
}
