function onMessageSendHandler(event) {
    event.completed({
        allowEvent: false,
        errorMessage: "TEST ROVECO MAC - EVENTO ONMESSAGESEND ATTIVATO"
    });
}

Office.onReady(function () {
    Office.actions.associate(
        "onMessageSendHandler",
        onMessageSendHandler
    );
});
