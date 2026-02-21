/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */
import { base64Image } from "../../base64Image.js";
import { processMessage, openDialog, tryCatch, insertIcon,
    insertParagraph, getSelectionText, insertStageDiagram } from './taskpaneFunctions.js';

// export var dialog = 'Hello from another file!';
// console.log("Set dialog to null");
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // Assign event handlers and other initialization logic.
    // document.getElementById("insert-paragraph").onclick = () => tryCatch(insertParagraph);
    debugger;
    document.getElementById("open-stage-window").onclick = () => tryCatch(openDialog);
    document.getElementById("get-characters").onclick = () => tryCatch(getCharacters);
    document.getElementById("get-speaker").onclick = () => tryCatch(getSpeaker);
    document.getElementById("insert-content-control").onclick = () => tryCatch(insertContentControl);
    document.getElementById("get-content-control").onclick = () => tryCatch(getContentControl);
    document.getElementById ("insert-stage-diagram").onclick = () => tryCatch(insertStageDiagram);
/*     document.getElementById("apply-style").onclick = () => tryCatch(applyStyle);
    document.getElementById("apply-custom-style").onclick = () => tryCatch(applyCustomStyle);
    document.getElementById("change-font").onclick = () => tryCatch(changeFont);
    document.getElementById("insert-text-into-range").onclick = () => tryCatch(insertTextIntoRange);
    document.getElementById("insert-text-outside-range").onclick = () => tryCatch(insertTextBeforeRange);
    document.getElementById("replace-text").onclick = () => tryCatch(replaceText);
    document.getElementById("insert-image").onclick = () => tryCatch(insertImage);
    document.getElementById("insert-html").onclick = () => tryCatch(insertHTML);
    document.getElementById("insert-table").onclick = () => tryCatch(insertTable);
    document.getElementById("create-content-control").onclick = () => tryCatch(createContentControl);
    document.getElementById("replace-content-in-control").onclick = () => tryCatch(replaceContentInControl); */
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
  }
});


/* async function insertParagraph() {
    await Word.run(async (context) => {

        // TODO1: Queue commands to insert a paragraph into the document.
        const docBody = context.document.body;
        docBody.insertParagraph("Office has several versions, including Office 2021, Microsoft 365 subscription, and Office on the web.",
                    Word.InsertLocation.start);
        await context.sync();
    });
} */



async function getCharacters(){
    await Word.run(async (context) => {
        const docBody = context.document.body;
        const paragraphs = docBody.paragraphs.load('style, text');

        const targetStyleName = "Speaker";
        const foundParagraphs = [];
        await context.sync();
        // Loop through the paragraphs to find those with the specific style
        for (let i = 0; i < paragraphs.items.length; i++) {
            // The style property returns the style name
            if (paragraphs.items[i].style === targetStyleName) {
                foundParagraphs.push(paragraphs.items[i]);
            }
        }
        var x = foundParagraphs;
    });
}
// Random change to permit me to push to github

/** Default helper for invoking an action and handling errors. */
/* async function tryCatch(callback) {
    try {
        await callback();
    } catch (error) {
        // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
        console.error(error);
    }
}
 */
async function applyStyle() {
    await Word.run(async (context) => {

        // TODO1: Queue commands to style text.
        const firstParagraph = context.document.body.paragraphs.getFirst();
        firstParagraph.styleBuiltIn = Word.BuiltInStyleName.intenseReference;
        await context.sync();
    });
}

async function applyCustomStyle() {
    await Word.run(async (context) => {

        // TODO1: Queue commands to apply the custom style.
        const lastParagraph = context.document.body.paragraphs.getLast();
        lastParagraph.style = "MyCustomStyle";
        await context.sync();
    });
}

async function changeFont() {
    await Word.run(async (context) => {

        // TODO1: Queue commands to apply a different font.
        const secondParagraph = context.document.body.paragraphs.getFirst().getNext();
        secondParagraph.font.set({
        name: "Courier New",
        bold: true,
        size: 18
    });
        await context.sync();
    });
}

async function insertTextIntoRange() {
    await Word.run(async (context) => {

        // TODO1: Queue commands to insert text into a selected range.
        const doc = context.document;
        const originalRange = doc.getSelection();
        originalRange.insertText(" (M365)", Word.InsertLocation.end);
        // TODO2: Load the text of the range and sync so that the
        //        current range text can be read.
        originalRange.load("text");
        await context.sync();
        // TODO3: Queue commands to repeat the text of the original
        //        range at the end of the document.
        doc.body.insertParagraph("Original range: " + originalRange.text, Word.InsertLocation.end);
        await context.sync();
    });
}

async function insertTextBeforeRange() {
    await Word.run(async (context) => {

        // TODO1: Queue commands to insert a new range before the
        //        selected range.
        const doc = context.document;
        const originalRange = doc.getSelection();
        originalRange.insertText("Office 2024, ", Word.InsertLocation.before);
        // TODO2: Load the text of the original range and sync so that the
        //        range text can be read and inserted.
        originalRange.load("text");
        await context.sync();

        // TODO3: Queue commands to insert the original range as a
        //        paragraph at the end of the document.
        doc.body.insertParagraph("Current text of original range: " + originalRange.text, Word.InsertLocation.end);
        // TODO4: Make a final call of context.sync here and ensure
        //        that it runs after the insertParagraph has been queued.
        await context.sync();
    });
}

async function replaceText() {
    await Word.run(async (context) => {

        // TODO1: Queue commands to replace the text.
        const doc = context.document;
        const originalRange = doc.getSelection();
        originalRange.insertText("many", Word.InsertLocation.replace);
        await context.sync();
    });
}

async function insertImage() {
    await Word.run(async (context) => {

        // TODO1: Queue commands to insert an image.
        context.document.body.insertInlinePictureFromBase64(base64Image, Word.InsertLocation.end);
        await context.sync();
    });
}

async function insertHTML() {
    await Word.run(async (context) => {

        // TODO1: Queue commands to insert a string of HTML.
        const blankParagraph = context.document.body.paragraphs.getLast().insertParagraph("", Word.InsertLocation.after);
        blankParagraph.insertHtml('<p style="font-family: verdana;">Inserted HTML.</p><p>Another paragraph</p>', Word.InsertLocation.end);
        await context.sync();
    });
}

async function insertTable() {
    await Word.run(async (context) => {

        // TODO1: Queue commands to get a reference to the paragraph
        //        that will precede the table.
        const secondParagraph = context.document.body.paragraphs.getFirst().getNext();
        // TODO2: Queue commands to create a table and populate it with data.
        const tableData = [
          ["Name", "ID", "Birth City"],
          ["Bob", "434", "Chicago"],
          ["Sue", "719", "Havana"],
        ];
        secondParagraph.insertTable(3, 3, Word.InsertLocation.after, tableData);
        await context.sync();
    });
}

async function createContentControl() {
    await Word.run(async (context) => {
        debugger;
        // TODO1: Queue commands to create a content control.
        // User must first select the text that is to become the content control
        const serviceNameRange = context.document.getSelection();
        const serviceNameContentControl = serviceNameRange.insertContentControl();
        serviceNameContentControl.title = "Service Name";
        serviceNameContentControl.tag = "serviceName";
        serviceNameContentControl.appearance = "Tags";
        serviceNameContentControl.color = "blue";
        await context.sync();
    });
}

async function replaceContentInControl() {
    await Word.run(async (context) => {

        // TODO1: Queue commands to replace the text in the Service Name
        //        content control.
        const serviceNameContentControl = context.document.contentControls.getByTag("serviceName").getFirst();
        serviceNameContentControl.insertText("Fabrikam Online Productivity Suite", Word.InsertLocation.replace);
        await context.sync();
    });
}

async function getSpeakerName(){
    return await Word.run(async (context) => {
        const doc = context.document;
        const originalRange = doc.getSelection();
        originalRange.load("paragraphs, style");
        await context.sync();
        var thisPara = originalRange.paragraphs.items[0];
        if (thisPara.style == "Speaker"){
            thisPara.select();
            var charName = thisPara.text;
            // console.log("The selected para uses the Speaker style." + charName);
            return charName;
        }
        else {
            // Look back up the paras until we find a Speaker style
            while (true){
                thisPara = thisPara.getPreviousOrNullObject();
                thisPara.load("text, style");
                await context.sync();
                if(thisPara.isNullObject){
                    // console.log("No previous paragraph uses the Speaker style.")
                    return "Nobody";
                } 
                if(thisPara.style == "Speaker"){
                    // console.log("Found a Speaker style. The text is " + thisPara.text);
                    return thisPara.text;
                }
            }
        };
    });
}

async function getSpeaker() {
    // debugger;
    // console.log("In getSpeaker -- about to await getSpeakerName.")
    const speakerName = await getSpeakerName();
    // console.log("The speaker's name is " + speakerName);
}


async function insertContentControl() {
    await Word.run(async (context) => {
        const doc = context.document;
        const originalRange = doc.getSelection();
        const contentControl = originalRange.insertContentControl("PlainText");
        contentControl.appearance = "Tags";
        // Customize the properties of the new content control (optional)
           /* contentControl.title = "MyCustomControl";
            contentControl.tag = "unique_tag_123";
            contentControl.placeholderText = "Enter information here";
            contentControl.cannotDelete = false; */
        await context.sync();
    });
}

async function getContentControl () {
    await Word.run(async (context) => {
        var contentControl = await SelectedContentControl();
        var speakerName = await  getSpeakerName();
        if(contentControl){
            console.log ("Selected content control ID is " + contentControl.id + " with text: " + contentControl.text);
            console.log ("The speaker who is moving is: " + speakerName);
        }
        else {
            console.log ("No content control selected.");
        }
    });
}

async function SelectedContentControl() {
    return await Word.run(async (context) => {
        const doc = context.document;
        const selectedRange = doc.getSelection();
        // Get the parent content control of the selection. 
        // Use the *OrNullObject method, which returns an object with isNullObject
        // property set to true if no parent CC exists, instead of throwing an error.
        const parentContentControl = selectedRange.parentContentControlOrNullObject;
        // Load the 'id' or any other property of the parentContentControl to check its existence
    // 'id' is a good property because every content control has one.
    parentContentControl.load('id, text');

    // Synchronize the document state by executing the queued commands
    await context.sync();

    // Check if the isNullObject property is true
    if (parentContentControl.isNullObject) {
        console.log("The selection is not in a content control.");
        return null;
        // Perform actions when the selection is outside a content control
    } else {
        console.log(`The selection is inside content control with ID: ${parentContentControl.id}`);
        // Perform actions when the selection is inside a content control
        return parentContentControl;
    }
    });
}

let stageWindow;

/* function openStageWindow() {
  Office.context.ui.displayDialogAsync(
    "https://localhost:3000/stagewindow.html", // URL of the page to open in the dialog
    { height: 50, width: 50 }, // Options for size (percentage of the current window)
    (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log(asyncResult.error.code + ": " + asyncResult.error.message);
      } else {
        stageWindow = asyncResult.value;
        // Add event handlers, e.g., for messages or closure events
        // stageWindow.addEventHandler(Office.EventType.DialogEventReceived, processDialogMessage);
        stageWindow.addEventHandler(Office.EventType.DialogMessageReceived, processDialogMessage);
      }
    }
  );
} */

// Function to process messages received from the dialog
/* function processDialogMessage(arg) {
    debugger;
    const messageFromDialog = JSON.parse(arg.message);
    // Use the message data in your task pane
    console.log("Received data from dialog:", messageFromDialog);
} */