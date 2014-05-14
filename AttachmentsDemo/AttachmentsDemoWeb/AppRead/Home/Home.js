/// <reference path="../App.js" />

(function () {
    "use strict";

    // The Office initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            $('#getAttachment').click(getAttachment);
            $('#saveAttachment').click(saveAttachment);
            $('#providePermission').click(doOAuthFlow);

            displayItemDetails();
        });
    };

    // Displays the "Subject" and "From" fields, based on the current mail item
    function displayItemDetails() {
        var item = Office.context.mailbox.item;
        $('#subject').text(item.subject);

        var from;
        if (item.itemType === Office.MailboxEnums.ItemType.Message) {
            from = item.from;
        } else if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
            from = item.organizer;
        }

        if (from) {
            $('#from').text(from.displayName);
            $('#from').click(function () {
                app.showNotification(from.displayName, from.emailAddress);
            });
        }
    }

    function doOAuthFlow() {
        var dataToSend = {
            //ResourceId : "https://oauthplay-my.sharepoint.com/"
            ResourceId : "https://microsoft705-my.sharepoint.com"
        }

        $.ajax({
            url: '../../api/OAuth/GetAuthorizationUrl',
            type: 'POST',
            data: JSON.stringify(dataToSend),
            contentType: 'application/json;charset=utf-8'
        }).done(function (data) {
            // the data returned IS the url, so just open the window
            window.open(data);
        }).fail(function (status) {
            app.showNotification('Error', JSON.stringify(status));
        }).always(function () {
            $('.disable-while-sending').prop('disabled', false);
        });
    }

    function saveAttachment() {
        $('.disable-while-sending').prop('disabled', true);

        var attachmentId = Office.context.mailbox.item.attachments[0].id;
        var ewsUrl = Office.context.mailbox.ewsUrl;
        Office.context.mailbox.getCallbackTokenAsync(function (ar) {
            var attachmentData = {
                AuthToken: ar.value,
                AttachmentId: attachmentId,
                EwsUrl: ewsUrl
            };

            sendRequest("GetAttachment/SaveAttachment", attachmentData);
        });
    }
    
    function getAttachment() {
        $('.disable-while-sending').prop('disabled', true);

        var attachmentId = Office.context.mailbox.item.attachments[0].id;
        var ewsUrl = Office.context.mailbox.ewsUrl;
        Office.context.mailbox.getCallbackTokenAsync(function (ar) {
            var attachmentData = {
                AuthToken: ar.value,
                AttachmentId: attachmentId,
                EwsUrl: ewsUrl
            };

            sendRequest("GetAttachment/GetAttachment", attachmentData);
        });
    }

    // Helper method
    function sendRequest(method, data) {
        $.ajax({
            url: '../../api/' + method,
            type: 'POST',
            data: JSON.stringify(data),
            contentType: 'application/json;charset=utf-8'
        }).done(function (data) {
            app.showNotification("Success", JSON.stringify(data));
            console.log(JSON.stringify(data));
        }).fail(function (status) {
            app.showNotification('Error', JSON.stringify(status));
        }).always(function () {
            $('.disable-while-sending').prop('disabled', false);
        });
    }
})();