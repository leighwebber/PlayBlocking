

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
        // Assign event handlers and other initialization logic.
        // document.getElementById("insert-paragraph").onclick = () => tryCatch(insertParagraph);
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

        function iconCreate(initials, colour, state){
            icon.initials = initials;
            icon.colour = colour;
            icon.state = state;
            return icon;
        }
    }
});

export {iconCreate}