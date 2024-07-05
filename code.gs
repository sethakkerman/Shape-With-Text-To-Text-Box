function adjustTextBoxes() {
  console.log('Function Started');
  var presentation = SlidesApp.getActivePresentation();
  var slides = presentation.getSlides();
  console.log('Number of slides: ' + slides.length);

  slides.forEach(function(slide, index) {
    console.log('Processing slide: ' + (index + 1));
    var elements = slide.getPageElements();
    console.log('Number of elements in slide: ' + elements.length);
    elements.forEach(function(element, idx) {
      console.log('Processing element: ' + (idx + 1));
      if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
        var shape = element.asShape();
        if (shape.getShapeType() === SlidesApp.ShapeType.RECTANGLE && shape.getText().asString().trim() !== '') {
          console.log('Shape is a rectangle with text');
          
          var text = shape.getText().asString();
          var width = shape.getWidth();
          var height = shape.getHeight();
          var transform = shape.getTransform();
          var originalX = transform.getTranslateX();
          var originalY = transform.getTranslateY();
          console.log(`Original position: X=${originalX}, Y=${originalY}`);
          console.log(`Original size: Width=${width}, Height=${height}`);

          // Adjust for 0.2 inches padding by adding 14.4 points
          var paddedWidth = width + 14.4;
          var paddedHeight = height + 14.4;
          var paddedX = originalX - 7.2; // Shift left to compensate for increased width
          var paddedY = originalY - 7.2; // Shift up to compensate for increased height

          var textBox = slide.insertTextBox(text, paddedX, paddedY, paddedWidth, paddedHeight);

          // Apply formatting from original shape
          var textStyle = shape.getText().getTextStyle();
          var textBoxStyle = textBox.getText().getTextStyle();
          textBoxStyle.setBold(textStyle.isBold());
          textBoxStyle.setItalic(textStyle.isItalic());
          textBoxStyle.setFontSize(textStyle.getFontSize());
          textBoxStyle.setForegroundColor(textStyle.getForegroundColor());

          // Optionally, delete the original shape
          element.remove();
          console.log('New text box created with original text and formatting');
        } else {
          console.log('Element is not a text-containing rectangle, it is a ' + shape.getShapeType());
        }
      } else {
        console.log('Element is not a shape, it is a ' + element.getPageElementType());
      }
    });
  });
  console.log('Function Completed');
}

function onOpen() {
  var ui = SlidesApp.getUi();
  ui.createMenu('Custom Scripts')
    .addItem('Adjust Text Boxes', 'adjustTextBoxes')
    .addToUi();
}
