Office.onReady(() => {
    document.getElementById("disableLinks").addEventListener("click", disableLinks);
});

function disableLinks() {
    Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            let emailBody = result.value;
            
            // Replace all <a> tags with disabled text
            let updatedBody = emailBody.replace(/<a\b[^>]*>(.*?)<\/a>/gi, '<span style="color: red;">🔒 Link Disabled</span>');

            // Update the email body
            Office.context.mailbox.item.body.setAsync(updatedBody, { coercionType: Office.CoercionType.Html }, function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    console.log("Links disabled successfully.");
                } else {
                    console.error("Error disabling links: " + result.error.message);
                }
            });
        }
    });
}