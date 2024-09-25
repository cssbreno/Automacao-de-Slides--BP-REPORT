// Listagem de imagens no slide com ID, índice, largura e altura
function listImages() {
    var SLIDES_ID = '1whKu0ncTiW3qFmPUXusqTL7zu_ai_uy4kz09oyoaq34';
    var slideIndex = 8; // Índice do slide que deseja listar as imagens (0 é o primeiro slide)
    var SLIDES = SlidesApp.openById(SLIDES_ID);
    var slide = SLIDES.getSlides()[slideIndex];
    var images = slide.getImages();
    
    for (var i = 0; i < images.length; i++) {
      var image = images[i];
      var imageId = image.getObjectId(); // Obtendo o ID da imagem
      var width = image.getWidth(); // Obtendo a largura da imagem
      var height = image.getHeight(); // Obtendo a altura da imagem
      Logger.log('Image Index: ' + i + '\nImage ID: ' + imageId + '\nWidth: ' + width + '\nHeight: ' + height);
    }
  }
  