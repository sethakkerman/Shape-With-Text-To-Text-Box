function adjustTextBoxes() {
  console.log('Function Started');
  try {
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
          console.log('Element is a shape');
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
            if (textStyle.getFontSize()) {
              textBoxStyle.setFontSize(textStyle.getFontSize());
            }
            try {
              if (textStyle.getForegroundColor()) {
                textBoxStyle.setForegroundColor(textStyle.getForegroundColor());
              }
            } catch (e) {
              console.error('Error setting foreground color: ' + e.message);
            }

            // Apply paragraph and line spacing
            var originalTextRange = shape.getText();
            var newTextRange = textBox.getText();
            var originalParagraphs = originalTextRange.getParagraphs();
            var newParagraphs = newTextRange.getParagraphs();

            for (var i = 0; i < originalParagraphs.length; i++) {
              var originalParagraph = originalParagraphs[i];
              var newParagraph = newParagraphs[i];

              var originalParagraphStyle = originalParagraph.getRange().getParagraphStyle();
              var newParagraphStyle = newParagraph.getRange().getParagraphStyle();

              newParagraphStyle.setLineSpacing(originalParagraphStyle.getLineSpacing());
              newParagraphStyle.setSpaceAbove(originalParagraphStyle.getSpaceAbove());
              newParagraphStyle.setSpaceBelow(originalParagraphStyle.getSpaceBelow());
            }

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
  } catch (error) {
    console.error('Error: ' + error.message);
  }
}

function onOpen() {
  var ui = SlidesApp.getUi();
  ui.createMenu('Custom Scripts')
    .addItem('Adjust Text Boxes', 'adjustTextBoxes')
    .addToUi();
}
