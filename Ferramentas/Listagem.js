   // Listagem de formas no slide
   function listForms() {
    var SLIDES_ID = '1whKu0ncTiW3qFmPUXusqTL7zu_ai_uy4kz09oyoaq34';
    var slideIndex = 2; // Índice do slide que deseja listar os shapes (0 é o primeiro slide)
    var SLIDES = SlidesApp.openById(SLIDES_ID);
    var slide = SLIDES.getSlides()[slideIndex];
    var shapes = slide.getShapes();
    
    for (var i = 0; i < shapes.length; i++) {
      var shape = shapes[i];
      Logger.log('Shape ' + i + ': ' + shape.getText().asString());
    }
  }