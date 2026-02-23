var stageImageSource;
var stageImage;
var fromLoad = false;

// import { iconCreate } from "../movements/movements.js"; 

Office.onReady((info) => {
    Office.context.ui.addHandlerAsync(
        Office.EventType.DialogParentMessageReceived,
        onMessageFromParent,
        onRegisterMessageComplete
    );
    debugger;
    document.getElementById("testButton").onclick = () => tryCatch(testButton);
    fromLoad = true;
    window.addEventListener('resize', (event) => {
        // debugger;
        drawStageImage();
    });
    fromLoad = false;
});
function testButton(){
    debugger;
    const flexContainer = document.getElementById ("flex-container");
    const flexPanelUpper = document.getElementById("flex-panel-upper");
    const flexPanelLower = document.getElementById("flex-panel-lower");
    const stageImage = document.getElementById("stage-image");
    const infoDiv = document.getElementById("info");

    infoDiv.innerText += "flexPanelUpper.clientHeight: " + flexPanelUpper.clientHeight +
        "  window height: " + window.innerHeight + 
        "  flexPanelUpper.clientWidth: " + flexPanelUpper.clientWidth +
        "  window width: " + window.innerWidth + "\n";
    var icon = iconCreate("WG", 25, "blue", "normal");
    infoDiv.innerText += "Icon initials: " + icon.initials + "\n";

    }
function sendStringToParentPage() {
    const userName = document.getElementById("name-box").value;
    Office.context.ui.messageParent(userName);
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
/**
  * Conserve aspect ratio of the original region. Useful when shrinking/enlarging
  * images to fit into a certain area.
  *
  * @param {Number} srcWidth width of source image
  * @param {Number} srcHeight height of source image
  * @param {Number} maxWidth maximum available width
  * @param {Number} maxHeight maximum available height
  * @return {Object} { width, height }
  */
function calculateAspectRatioFit(widthNative, heightNative, widthWindow, heightWindow) {
    const ratioNative = heightNative / widthNative;
    const ratioWindow = heightWindow / widthWindow;
    var heightCanvas;
    var widthCanvas;
    if (ratioWindow < ratioNative){
        heightCanvas = heightWindow;
        widthCanvas = heightWindow / ratioNative;
    }
    else {
        heightCanvas = widthWindow * ratioNative;
        widthCanvas = widthWindow;
    }
    return({width: widthCanvas, height: heightCanvas});
 }
 function onMessageFromParent(arg) {
    // debugger;
    const messageFromParent = JSON.parse(arg.message);
    document.getElementById("testButton").onclick = () => tryCatch(testButton);
    switch(messageFromParent.messageType){
        case "sillyStuff":
            document.getElementById("message_text").innerText = messageFromParent.text;
            break;
        case "iconPlace":
            const newButton = document.createElement("button");
            newButton.innerText = "LW";
            newButton.id = "iconLW";
            newButton.type = 'button';
            newButton.addEventListener('click', () => {
                console.log('You clicked ' + newButton.innerText);
            });
            newButton.style.top = "100px";
            newButton.style.left = "50px";
            newButton.style.width = "40px";
            newButton.style.height = "20px";
            document.body.appendChild(newButton);
            break;
        case "imageLoad":
            /* stageImage = new Image();
            stageImage.onload = function(){
                drawStageImage();
            }
            stageImageSource = messageFromParent.src;
            stageImage.src = stageImageSource;
            drawStageImage(); */
            debugger;

            var img = new Image();
            img.id = "stage-image";
            img.className = "stageImage";
            img.alt = "Stage image";
            img.onload = function() {
                drawStageImage(img); 
            }
            img.src = messageFromParent.src;
            // var stageImage = document.getElementById("stage-image");
            // stageImage.src = messageFromParent.src;
            
            // drawStageImage(stageImage);
            break;
        };
}


function drawStageImage(stageImage){
    // if(fromLoad) return;
    // if(stageImage){
        debugger;
        const flexContainer = document.getElementById("flex-container");
        const flexPanelUpper = document.getElementById("flex-panel-upper");
        const flexPanelLower = document.getElementById("flex-panel-lower");

        // var stageImageElement = document.getElementById("stage-image");
        // stageImage.src = src;

        // const stageImage = document.getElementById("stage-image");
        // const ctx = myCanvas.getContext("2d");
        
        // myCanvas.width = window.innerWidth;
        // myCanvas.width = flexPanelUpper.clientWidth;
        stageImage.width = flexPanelUpper.clientWidth;
        // myCanvas.height = window.innerHeight - 100;
        // myCanvas.height = flexPanelUpper.clientHeight;
        stageImage.height = flexPanelUpper.clientHeight;
        

        var naturalWidth = stageImage.naturalWidth;
        var naturalHeight = stageImage.naturalHeight;
        var newSize = calculateAspectRatioFit(naturalWidth, naturalHeight, 
            flexPanelUpper.clientWidth, flexPanelUpper.clientHeight);
        // debugger;
        // ctx.clearRect(0, 0, myCanvas.width, myCanvas.height);
        // myCanvas.width = newSize.width;
        stageImage.width = newSize.width;
        // myCanvas.height = newSize.height;
        stageImage.height = newSize.height;

        /* ctx.drawImage(stageImage, 0, 0, 
            newSize.width, newSize.height); */
        flexPanelUpper.clientHeight = newSize.height;
        flexPanelUpper.clientWidth = newSize.width;
        flexPanelUpper.append(stageImage);
    // };
}
function onRegisterMessageComplete(asyncResult) {
    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        // reportError(asyncResult.error.message);
        console.log("ERROR: " + asyncResult.error.message);
    }
}

const icon = {};
Object.defineProperty(icon, "initials", {
value: "--",
writable: true,
enumerable: true,
configurable: false,   
});
Object.defineProperty(icon, "colour", {
    value: "black",
    writable: true,
    enumerable: true,
    configurable: false,   
});
// An icon's state can be normal, disabled, or ghosted
Object.defineProperty(icon, "state", {
    value: "normal",
    writable: true,
    enumerable: true,
    configurable: false,   
});
// diameter in pixels
Object.defineProperty(icon, "diameter", {
    value: 20,
    writable: true,
    enumerable: true,
    configurable: false,   
});

function iconCreate(initials, diameter, colour, state){
    icon.initials = initials;
    icon.diameter = diameter;
    icon.colour = colour;
    icon.state = state;
    return icon;
}
