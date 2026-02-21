(function () {
    Office.initialize = function (reason) {
        // Office is ready
        $(document).ready(function () {
            document.getElementById("send-message").onclick = () => tryCatch(sendDataToTaskPane);
        })
        // document.getElementById("send-message").onclick = () => tryCatch(sendDataToTaskPane);
    };

    

    function sendDataToTaskPane() {
        debugger;
        const dataToPass = {
            setting1: "value1",
            setting2: "value2"
        };

        // Convert the data to a JSON string for transmission
        const message = JSON.stringify(dataToPass);

        // Send the message to the host page (task pane)
        Office.context.ui.messageParent(message);
    }

});

