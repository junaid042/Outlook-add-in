Office.onReady((info) => {
  // Office is ready
  if (info.host === Office.HostType.Outlook) {
  }
});

let token; // Ensure that 'token' is defined globally
let fromEmail; 
var messageId;



function getItemRestId() {
  if (Office.context.mailbox.diagnostics.hostName === 'OutlookIOS') {
    // itemId is already REST-formatted.
    return Office.context.mailbox.item.itemId;
  } else {
    // Convert to an item ID for API v2.0.
    return Office.context.mailbox.convertToRestId(
      Office.context.mailbox.item.itemId,
      Office.MailboxEnums.RestVersion.v2_0
    );
  }
}

function uploadWholeEml() {
  Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function (result) {
    var ewsId = Office.context.mailbox.item.itemId;
    var tokenA = result.value;
    var item = Office.context.mailbox.item;

    var getMessageUrl = Office.context.mailbox.restUrl +
      '/v2.0/me/messages/' + getItemRestId() + '/$value';

    fetch(getMessageUrl, {
      method: 'GET',
      headers: new Headers({
        Authorization: 'Bearer ' + tokenA
      })
    }).then(function (response) {
      if (!response.ok) {
        throw new Error("Error fetching message: " + response.status);
      }
      return response.blob();
    }).then(function (blob) {
      var reader = new FileReader();
      reader.readAsDataURL(blob);
      reader.onloadend = function () {
        var base64data = reader.result;
        base64 = base64data.split(',')[1];
        run( base64);
      };
    }).catch(function (error) {
      console.error("Fetch error:", error);
    });
  });
}

function spamEmail(event) {
  var a = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=fdbb3010-2e84-4073-8b95-712c5861f36a&response_type=token&redirect_uri=https://junaid042.github.io/Outlook-add-in/assets/Redirect.html&scope=user.read%20mail.readwrite%20mail.send&response_mode=fragment&state=12345&nonce=678910";
  Office.context.ui.displayDialogAsync(a, {  height: 60, width: 40 }, function (asyncResult) {
    Logindialog = asyncResult.value;
    Logindialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
      token = arg.message;
      uploadWholeEml();
      Logindialog.close();
    });
  });
}

function run(base64) {
  var item = Office.context.mailbox.item;

  if (item.itemType === Office.MailboxEnums.ItemType.Message) {
    fromEmail = item.sender;
    console.log("Sender Email Address: " + fromEmail.emailAddress);
  } else {
    console.log("This is not an email message.");
  }
 
  const userProfile = Office.context.mailbox.userProfile;
  console.log(`Hello ${userProfile.emailAddress}`);
  const sender = userProfile.emailAddress

    var emailSettings = {
      url: "https://graph.microsoft.com/v1.0/me/sendMail",
      method: "POST",
      timeout: 0,
      headers: {
        "Content-Type": "application/json",
        Authorization: "Bearer " + token,
      },
      data: JSON.stringify({
        message: {
          subject: "Email Reported as SPAM",
          body: {
            contentType: "text",
            content: "Sender: " + sender + "\nSuspicious Address: " + fromEmail.emailAddress,
          },
          toRecipients: [
            {
              emailAddress: {
                address: sender, ////for now add-in user recive this email you can change this where you want  
              },
            },
          ],
          attachments: [
            {
              "@odata.type": "#microsoft.graph.fileAttachment",
              name: "email.eml",
              contentType: "message/rfc822",
              contentBytes: base64,
            },
          ],
        },
        saveToSentItems: "false",
      }),
    };

    return fetch(emailSettings.url, {
      method: emailSettings.method,
      headers: emailSettings.headers,
      body: emailSettings.data,
    }).then((response) => {
    if (!response.ok) {
      throw new Error("Error sending email: " + response.status);
    }
    return response;
  }).then(()=>{
    deleteEmail(token);
  }).catch((error) => {
    console.error("Error:", error);
  });
}

function deleteEmail(token){
  messageId = Office.context.mailbox.item.itemId;
  messageId = messageId.replaceAll("/", "-");
  
  var settings = {
    url: "https://graph.microsoft.com/v1.0/me/messages/" + messageId,
    method: "DELETE",
    timeout: 0,
    headers: {
      Authorization: "Bearer " + token,
    },
  };
  $.ajax(settings).done(function (response) {
    console.log(response);
  });
  }

Office.actions.associate("spamEmail", spamEmail);
