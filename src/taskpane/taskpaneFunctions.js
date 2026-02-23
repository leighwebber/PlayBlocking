var dialog = null;
var dialogOpened = false;

async function insertParagraph() {
    await Word.run(async (context) => {

        // TODO1: Queue commands to insert a paragraph into the document.
        const docBody = context.document.body;
        docBody.insertParagraph("The VERY SLOW brown fox jumps over the lazy dog.",
            Word.InsertLocation.start);
        await context.sync();
        
    });
}
async function insertIcon() {
    await Word.run(async (context) => {
        if(dialog){
            /* const messageObject = { messageType: "sillyStuff", text: "Hello there"};
            var jsonMessage = JSON.stringify(messageObject);
            dialog.messageChild(jsonMessage); */
            const iconObject = { messageType: "iconPlace", xPct: 0.30, yPct: 0.25, width: 40, height: 40};
            const jsonMessage = JSON.stringify(iconObject);
            dialog.messageChild(jsonMessage);
        }
    });
}

function openDialog() {
    Office.context.ui.displayDialogAsync("https://localhost:3000/popup.html", { height: 45, width: 55 },
    (asyncResult) => {
        dialog = asyncResult.value;
        dialogOpened = true;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
            // dialog.close();
            processMessage(arg);
        });
    });
}
function processMessage(arg) {
    debugger;
    var message = JSON.parse(arg);
    switch(message.messageType){
        case "userName":
            console.log(arg.userName);
            break;
        case "popupData":
            var popupDataJason = JSON.stringify(arg.popupData);
            console.log(popupDataJason);
            break;
    }
}

function onRegisterMessageComplete(asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log(asyncResult.error.message);
        return;
    }
}
/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
    try {
        await callback();
    } catch (error) {
        // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
        console.error(error);
    }
}
async function insertStageDiagram() {
    await Word.run(async (context) => {
        // debugger;
        if(dialogOpened){
            /* const messageObject = { messageType: "sillyStuff", text: "Hello there"};
            var jsonMessage = JSON.stringify(messageObject);
            dialog.messageChild(jsonMessage); */
            // debugger;
            const messageObject = { messageType: "imageLoad", src: "../../assets/AnneFrankSet.jpg"};
            const jsonMessage = JSON.stringify(messageObject);
            dialog.messageChild(jsonMessage);
        }
    });
}


async function getSelectionText() {
    await Word.run(async (context) => {
        // debugger;
        const docBody = context.document.body;
        const range = context.document.getSelection();
        range.load("text");
        await context.sync();
        console.log("Selected text: " + range.text);
        
    });
}
export { openDialog, tryCatch, insertIcon,
    insertParagraph, getSelectionText, insertStageDiagram }