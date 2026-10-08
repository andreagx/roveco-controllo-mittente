var CUSTOMER_SERVICE_ADDRESS = "customerservice@rovecogroup.com";

function normalizeAddress(value) {
    return typeof value === "string" ? value.trim().toLowerCase() : "";
}

function onMessageSendHandler(event) {
    var completed = false;
    var timer;
    function finish(options) {
        if (completed) return;
        completed = true;
        if (timer) clearTimeout(timer);
        event.completed(options);
    }
    function unableToVerify() {
        finish({
            allowEvent: false,
            errorMessage: "Impossibile verificare il mittente. Controlla il campo Da: per Customer Service usa " +
                CUSTOMER_SERVICE_ADDRESS + ". Puoi scegliere Invia comunque se il mittente è corretto."
        });
    }
    timer = setTimeout(unableToVerify, 3000);
    try {
        var mailbox = Office.context.mailbox;
        var personalAddress = normalizeAddress(mailbox.userProfile.emailAddress);
        mailbox.item.from.getAsync(function (result) {
            if (completed) return;
            try {
                if (result.status !== Office.AsyncResultStatus.Succeeded) {
                    unableToVerify();
                    return;
                }
                var sender = normalizeAddress(result.value && result.value.emailAddress);
                if (!sender) {
                    unableToVerify();
                } else if (sender === CUSTOMER_SERVICE_ADDRESS) {
                    finish({ allowEvent: true });
                } else if (!personalAddress) {
                    unableToVerify();
                } else if (sender === personalAddress) {
                    finish({
                        allowEvent: false,
                        errorMessage: "Stai inviando da " + sender +
                            ". Per Customer Service usa " + CUSTOMER_SERVICE_ADDRESS +
                            ". Scegli Non inviare per cambiare mittente o Invia comunque per confermare."
                    });
                } else {
                    finish({ allowEvent: true });
                }
            } catch (error) {
                unableToVerify();
            }
        });
    } catch (error) {
        unableToVerify();
    }
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
