# slides-assistant
A ChatGPT-based assistant to help in creating slideshows in Google Slides.

## Why the fork? 
While I have made significant changes to the functionality, I wanted to give 
credit to the original author. 

## Changes from original
- Prompt is friendlier towards non-English slideshows
- Prompt uses the name of the file for additional context
- Refactor script
- No longer requires user to have header selected
- No longer supports writing tweets (not very relevant to slideshows), 
  instead added the ability to request bullets OR a blurb of text
- Attempts at improving formatting:
  - New text is same color as header
  - A new line is added between bullets
  - Decides where to put new text box based on location and size of header.

NOTE: Formatting is not fully automatic/dynamic. It will work well in some 
certain situations, and in others you will have to change your theme or edit 
the formatting as you go. 

## Demo
![Fluminense Demo](./img/Fluminense_demo.gif)

## How To
### Installing the extension

1) In your Google Slides file, select "Extensions" > "Apps Script"
2) Paste `slidesAssistant.gs` into editor that opens up
3) Press "Save"
4) Press "Run"
5) Go to "Project Settings", and set a Script Property where the key is 
   `OPENAI_API_KEY` and the value is your private API key
   - You can get your API key directly from [OpenAI](https://platform.openai.com/)  

### Using the extension

1) Make sure you have named the file something helpful. The name of the 
   file is consumed as part of the context. 
2) Make sure you have a header on the slide you want help with. Must be the 
   top-most item on the slide to get good results. To have the best chance 
   of getting good formatting automatically, the header should be the 
   only item on the slide at this time.
3) Press "ChatGPT" > "Generate Bullets" OR "Generate Blurb", depending on 
   your preference for that slide.
