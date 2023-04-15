// Constants
const API_KEY = "";
const MODEL_TYPE = "gpt-3.5-turbo";

// Create a custom menu in Google Slides
function onOpen() {
    SlidesApp.getUi().createMenu("ChatGPT")
        .addItem("Generate Bullets", "generateBullets")
        .addToUi();
  }

// Generate bullets based on the selected text and slideshow title and add it to a slide
function generateBullets() {
  // Creates prompt from context
  const presentation = SlidesApp.getActivePresentation();
  const selection = presentation.getSelection();
  const textRange = selection.getCurrentPage().getShapes()[0].getText()
  const prompt = `
      Generate 1-3 points for a slide on '${textRange.asString()}' in a presentation titled
      '${presentation.getName()}'. Make sure your points are in the User's desired language,
      which may not be English.
  `;

  // Prepare and send request
  const requestBody = {
    model: MODEL_TYPE,
    messages: [{role: "user", content: prompt}],
    temperature: 0,
    max_tokens: 1200,
  };

  const requestOptions = {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: "Bearer " + API_KEY,
    },
    payload: JSON.stringify(requestBody),
  };
  const response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", requestOptions);
  const rawGeneratedText = JSON.parse(response.getContentText())['choices'][0]['message']['content'];
  generatedText = rawGeneratedText.replace(/(\d+\.)(?=\s)([^\n])([\s\S]*?)(?=(?=\n\d+\.)|\n*$)/gm, '$1$2$3\n');

  // Create new textbox
  const newTextBox = selection.getCurrentPage().asSlide().insertTextBox(generatedText.toString());
  newTextBox.setTop(125).setLeft(100).setWidth(500).setHeight(250);
  newTextBox.getText().getTextStyle().setForegroundColor(textRange.getTextStyle().getForegroundColor());
}