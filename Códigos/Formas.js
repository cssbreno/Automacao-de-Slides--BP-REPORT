
// IDs da apresentação e da planilha
const SLIDES_ID = '1whKu0ncTiW3qFmPUXusqTL7zu_ai_uy4kz09oyoaq34';
const SHEETS_ID = '1GZAPqryb7C3J_rF2GpyMAYkqj5itteDZn_MD6cDmB1Q';

// Nome da planilha
const SHEET_NAME = 'Base';

 const slideIndex = 2; // Índice do slide (baseado em 0)  
  
  // Obter altura e largura de formas do slide
  function logShapeSize() {
  // Obter a apresentação ativa
  var presentation = SlidesApp.getActivePresentation();
  
  // Selecionar o slide que contém a forma
  var slide = presentation.getSlides()[3];
  
  // Selecionar a primeira forma no slide
  var shape = slide.getShapes()[8];
  
  // Obter a largura e a altura da forma
  var width = shape.getWidth();
  var height = shape.getHeight();
  
  // Registrar a largura e a altura da forma
  Logger.log('Largura da forma: ' + width + ' pixels');
  Logger.log('Altura da forma: ' + height + ' pixels');
}