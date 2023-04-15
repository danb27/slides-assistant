# chatgpt-google-slides
A Google Apps Script to Assist in Creating Google Slides

## Why the fork? 
While I have made some changes to the functionality, I wanted to give credit 
to the original author. 

## Changes from original
- Prompt is friendlier towards non-English slideshows
- Prompt uses the name of the file for additional context
- Minor attempts at improving formatting:
  - New text is same color as header
  - A new line is added between bullets
  - New Text box is placed in a reasonable starting location

NOTE: Formatting is not fully automatic/dynamic. It will work well in some 
certain situations, and in others you will have to change your theme or edit 
the formatting as you go. 

## Demo
![Fluminense Demo](./img/Fluminense_demo.gif)

## How To
### Installing the extension

1) In your Google Slides file, select "Extensions" > "Apps Script"
2) Paste `slidesAssistant.gs` into editor that opens up
3) Add your own OpenAI API key to the script
4) Press "Save"
5) Press "Run"

### Using the extension

1) Make sure you have named the file something helpful. The name of the 
   file is consumed as part of the context. 
2) Select the header of a slide you want to write bullet points for. 
3) Press "ChatGPT" > "Generate Bullets"
