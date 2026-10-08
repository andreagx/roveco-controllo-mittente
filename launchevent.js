function onMessageSendHandler(event) {
    event.completed({
        allowEvent: false,
        errorMessage: "TEST ROVECO MAC - EVENTO ONMESSAGESEND ATTIVATO"
    });
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
