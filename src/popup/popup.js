var stageImageSource;
var stageImage;
var fromLoad = false;
var popupData;

// const popupData = new PopupData(null, null, null, null);

Office.onReady((info) => {
    Office.context.ui.addHandlerAsync(
        Office.EventType.DialogParentMessageReceived,
        onMessageFromParent,
        onRegisterMessageComplete
    );
    // debugger;
    popupData = new PopupData(null, null, null, []);
    document.getElementById("testButton").onclick = () => tryCatch(testButton);
    fromLoad = true;
    window.addEventListener('resize', (event) => {
        // debugger;
        const stageImage = document.getElementById("stage-image");
        drawStageImage(stageImage);
    });
    window.addEventListener("beforeunload", function(e) {
        // Perform actions like sending an analytics request
        console.log("Page is about to unload. Performing final actions.");
        popupData.saveToParent();
    });
    document.addEventListener("click", function(event) {
        const xPos = event.clientX; // X coordinate relative to the viewport
        const yPos = event.clientY; // Y coordinate relative to the viewport
        const stageImage = document.getElementById("stage-image");
        const xRel = (xPos - stageImage.x) / stageImage.width;
        const yRel = (yPos - stageImage.y) / stageImage.height;
        console.log("Mouse clicked at X:", xPos, "Y:", yPos, "stageImage.x", stageImage.x,
            "stageImage.y", stageImage.y, "stageImage.width", stageImage.width, 
            "stageImage.height", stageImage.height, "xRel", xRel, "yRel", yRel);
    });
    fromLoad = false;
});
function testButton(){
    // debugger;
    sendStringToParentPage("userName", "Leigh");
    const flexContainer = document.getElementById ("flex-container");
    const flexPanelUpper = document.getElementById("flex-panel-upper");
    const flexPanelLower = document.getElementById("flex-panel-lower");
    const stageImage = document.getElementById("stage-image");
    const infoDiv = document.getElementById("info");

    infoDiv.innerText += "flexPanelUpper.clientHeight: " + flexPanelUpper.clientHeight +
        "  window height: " + window.innerHeight + 
        "  flexPanelUpper.clientWidth: " + flexPanelUpper.clientWidth +
        "  window width: " + window.innerWidth + "\n";
    const icon = new Icon("WG", 20, "blue", "normal");
    popupData.iconCollection.push(icon);
    icon.place("Foo place");
    infoDiv.innerText += "Icon initials: " + icon.initials + "\n";

    }
function sendStringToParentPage(text) {
    // debugger;
    var message = {
        messageType: "userName",
        userName: text
    }
    var messageJson = JSON.stringify(message)
    Office.context.ui.messageParent(messageJson);
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
        case "popupData":
            popupData.flexPanelClientHeight = messageFromParent.popupData.flexPanelClientHeight;
            popupData.flexPanelClientWidth = messageFromParent.popupData.flexPanelClientWidth;
            popupData.iconCollection = messageFromParent.popupData.iconCollection;
            popupData.imageSrc = messageFromParent.popupData.imageSrc;
        case "imageLoad":
            /* stageImage = new Image();
            stageImage.onload = function(){
                drawStageImage();
            }
            stageImageSource = messageFromParent.src;
            stageImage.src = stageImageSource;
            drawStageImage(); */
            // debugger;

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
            popupData.image = img;
            // drawStageImage(stageImage);
            break;
        };
}

function drawStageImage(stageImage){
    // if(fromLoad) return;
    // if(stageImage){
        // debugger;
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
        popupData.image = stageImage;
    // };
}
function onRegisterMessageComplete(asyncResult) {
    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        // reportError(asyncResult.error.message);
        console.log("ERROR: " + asyncResult.error.message);
    }
}

/* const icon = {};
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
*/

/*
function iconCreate(initials, diameter, colour, state){
    const icon = {
        initials: initials,
        diameter: diameter,
        colour: colour,
        state: state,
        place: function(params){
            console.log(params);
        }
    }
    icon.initials = initials;
    icon.diameter = diameter;
    icon.colour = colour;
    icon.state = state;
    icon.place = function(params){
        console.log(params);
    }
    return icon;
} */
class Icon {
    constructor(initials, diameter, colour, state){
        this.initials = initials;
        this.diameter = diameter;
        this.colour = colour;
        this.state = state;
        this.xRel = 0;
        this.yRel = 0;
        this.iconCollection = [];
    }
    place(params) {
        console.log(params);
    }
}
class PopupData {
    constructor(image, flexPanelClientWidth, flexPanelClientHeight, iconCollection) {
        this.image = image;
        this.flexPanelClientWidth = flexPanelClientWidth;
        this.flexPanelClientHeight = flexPanelClientHeight;
        this.iconCollection = iconCollection;
    }
    saveToParent() {
        var message = {
            messageType: "popData",
            popupData: JSON.stringify(this)
        }
        Office.context.ui.messageParent(message);
    }

}
