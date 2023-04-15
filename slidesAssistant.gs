// Constants
const MODEL_TYPE = "gpt-3.5-turbo";
const PRESENTATION = SlidesApp.getActivePresentation();
const CURRENT_SLIDE = PRESENTATION.getSelection().getCurrentPage().asSlide();
const HEADER_TEXT_BOX = CURRENT_SLIDE.getShapes()[0];

// Creates a custom menu in Google Slides
function onOpen() {
  SlidesApp.getUi()
    .createMenu("ChatGPT")
    .addItem("Generate Bullets", "generateBullets")
    .addItem("Generate Blurb", "generateBlurb")
    .addToUi();
}

function generateBullets() {
  // Creates prompt from context
  const prompt = `
      Generate 1-3 points for a slide on '${HEADER_TEXT_BOX.getText().asString()}' in a presentation titled
      '${PRESENTATION.getName()}'. Make sure your points are in the User's language,
      which may not be English.
  `;
  return generateTextBox(prompt);
}

function generateBlurb() {
  // Creates prompt from context
  const prompt = `
      Generate a the contents of the main text box on a slide titled '${HEADER_TEXT_BOX.getText().asString()}'
      in a presentation titled '${PRESENTATION.getName()}'. Make sure your blurb is in the User's
      language, which may not be English.
  `;
  return generateTextBox(prompt);
}

// Executes a prompt and adds content to a new text box
function generateTextBox(prompt) {
  // Call the API
  const generatedText = callChatGptApi(prompt);

  // Adds generated text to slide
  addGeneratedTextToSlide(generatedText);
}

function callChatGptApi(prompt) {
  const requestBody = {
    model: MODEL_TYPE,
    messages: [{ role: "user", content: prompt }],
    max_tokens: 1200,
  };

  const requestOptions = {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: "Bearer " + getApiKey(),
    },
    payload: JSON.stringify(requestBody),
  };

  const response = UrlFetchApp.fetch(
    "https://api.openai.com/v1/chat/completions",
    requestOptions
  );
  const rawGeneratedText = JSON.parse(response.getContentText())["choices"][0][
    "message"
  ]["content"];
  return rawGeneratedText.replace(
    /(\d+\.)(?=\s)([^\n])([\s\S]*?)(?=(?=\n\d+\.)|\n*$)/gm,
    "$1$2$3\n"
  );
}

// Gets API key from Script Properties
function getApiKey() {
  return PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
}

function addGeneratedTextToSlide(generatedText) {
  // Create a new text box with the generated text
  const newTextBox = CURRENT_SLIDE.insertTextBox(generatedText.toString());

  // Get position and sizes of header
  const headerTextBoxLeft = HEADER_TEXT_BOX.getLeft();
  const headerTextBoxTop = HEADER_TEXT_BOX.getTop();
  const headerTextBoxWidth = HEADER_TEXT_BOX.getWidth();
  const headerTextBoxHeight = HEADER_TEXT_BOX.getHeight();

  // Apply header formatting to new text
  newTextBox
    .setTop(headerTextBoxTop + headerTextBoxHeight + 20)
    .setLeft(headerTextBoxLeft)
    .setWidth(headerTextBoxWidth)
    .setHeight(300 - headerTextBoxHeight);
  newTextBox
    .getText()
    .getTextStyle()
    .setForegroundColor(
      HEADER_TEXT_BOX.getText().getTextStyle().getForegroundColor()
    );
}
