
import { Card, RadioGroup, Radio, Button } from "@fluentui/react-components";
import { Send16Filled } from "@fluentui/react-icons";
import React, { useState, } from "react";

const questions = [
  '¿Has pulsado algún enlace?',
  '¿Has introducido algún código de usuario o contraseña en las pantallas a las que has llegado?',
  '¿Has notado algún comportamiento anómalo?',
  'En las últimas horas o días, ¿Has recibido correos electrónicos similares?',
  'En los últimos días, ¿has accedido a tu correo electrónico desde algún dispositivo personal o desde una red diferente a la de Nortegas?',
];
export default function ReportPishing() {
  const [responses, setResponses] = useState(Array(questions.length).fill('No'));
  const [riskLevel, setRiskLevel] = useState('');
  const [showResult, setShowResult] = useState(false);

  const handleRadioChange = (index, value) => {
    const newResponses = [...responses];
    newResponses[index] = value;
    setResponses(newResponses);
  };

  const calculateRiskLevel = () => {
    const shouldShowBajo = responses.every(response => response === 'No');
    if (shouldShowBajo) {
      return 'BAJO';
    }

    const shouldShowMedio1 = (responses[0] === 'Sí' && responses[1] === 'No' && responses[2] === 'No' && (responses[3] === 'Sí' || responses[3] === 'No') && (responses[4] === 'Sí' || responses[4] === 'No'));
    if (shouldShowMedio1) {
      return 'MEDIO';
    }

    const shouldShowMedio2 = (responses[0] === 'No' && responses[1] === 'No' && responses[2] === 'No' && responses[3] === 'No' && responses[4] === 'Sí');
    if (shouldShowMedio2) {
      return 'MEDIO';
    }

    const shouldShowMedio3 = (responses[0] === 'No' && responses[1] === 'No' && responses[2] === 'No' && responses[3] === 'Sí' && responses[4] === 'No');
    if (shouldShowMedio3) {
      return 'MEDIO';
    }

    const shouldShowAlto = (responses[0] === 'Sí' && responses[1] === 'Sí' && responses[2] === 'No' && (responses[3] === 'Sí' || responses[3] === 'No') && (responses[4] === 'Sí' || responses[4] === 'No'));
    if (shouldShowAlto) {
      return 'ALTO';
    }

    const shouldShowCritico = ((responses[0] === 'Sí' || responses[0] === 'No') && (responses[1] === 'Sí' || responses[1] === 'No') && responses[2] === 'Sí' && (responses[3] === 'Sí' || responses[3] === 'No') && (responses[4] === 'Sí' || responses[4] === 'No'));
    if (shouldShowCritico) {
      return 'CRITICO';
    }
    return '';
  };

  const handleSendClick = () => {
    const calculatedRiskLevel = calculateRiskLevel();
    setRiskLevel(calculatedRiskLevel);
    dailogopen();
  };
  var messageId = "";
  var token = "";
  var Logindialog;
  // var sender = "";
  var fromEmail = "";
  let base64 = "";
  // var emlFile = "";
  Office.onReady(() => {
    // If needed, Office.js is ready to be called.
  });

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
      // var ewsId = Office.context.mailbox.item.itemId;
      Office.context.item
      var token = result.value;
      // const currentTime = new Date();
      // const tokenData = {
      //   token: token,
      //   time: currentTime.getHours()+":"+currentTime.getMinutes() +":"+currentTime.getSeconds()
      // };
      // localStorage.setItem('tokenData', JSON.stringify(tokenData));
      
      // var item = Office.context.mailbox.item;

      var getMessageUrl = Office.context.mailbox.restUrl +
        '/v2.0/me/messages/' + getItemRestId() + '/$value';
        
      fetch(getMessageUrl, {
        method: 'GET',
        headers: new Headers({
          Authorization: 'Bearer ' + token
        })
      }).then(function (response) {
        response.blob().then(function (blob) {
          var reader = new FileReader();
          reader.readAsDataURL(blob);
          reader.onloadend = function () {
            var base64data = reader.result;
            base64 = base64data.split(',')[1];
            run(base64);
          };
        });
      });
    });
  };

  function dailogopen() {
    var a = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=fdbb3010-2e84-4073-8b95-712c5861f36a&response_type=token&redirect_uri=https://localhost:3000/assets/Redirect.html&scope=user.read%20mail.readwrite%20mail.send&response_mode=fragment&state=12345&nonce=678910";
    Office.context.ui.displayDialogAsync(a, { height: 60, width: 40 }, function (asyncResult) {
      Logindialog = asyncResult.value;
      Logindialog.addEventHandler(Office.EventType.DialogMessageReceived, LogprocessMessage);
    });
  };

  function LogprocessMessage(arg) {
    setShowResult(true);
    token = arg.message;
    uploadWholeEml();
    Logindialog.close();
  }

  var emailBody = `
  <h2>Phishing Email Report</h2>
  <table border="1">
    <tr>
      <th>Question Number</th>
      <th>Question</th>
      <th>Answer</th>
    </tr>
    ${questions.map((question, index) => `<tr><td>${index + 1}</td><td>${question.trim()}</td><td>${responses[index]}</td></tr>`).join('')}
  </table>`;

  // var item = Office.context.mailbox.item;
  messageId = Office.context.mailbox.item.itemId;
  messageId = messageId.replaceAll("/", "-");
  console.log(messageId);

  function run(base64) {

    var item = Office.context.mailbox.item;

    // Check if the item type is 'message'
    if (item.itemType === Office.MailboxEnums.ItemType.Message) {
      // Get the sender's email address
      fromEmail = item.sender;

      // Log the sender's email address to the console (you can modify this as needed)
      console.log("Sender Email Address: " + fromEmail.emailAddress);
    } else {
      // If the current item is not a message, handle accordingly
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
          subject: `Phishing Email Report  ${calculateRiskLevel()}`,
          body: {
            contentType: Office.CoercionType.Html,
            content: emailBody + "Sender: " + sender + " Suspicious Address: " + fromEmail.emailAddress,
          },
          toRecipients: [
            {
              emailAddress: {
                address: "junaid042@outlook.com",
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
    // Using Fetch API for the email sending part
    return fetch(emailSettings.url, {
      method: emailSettings.method,
      headers: emailSettings.headers,
      body: emailSettings.data,
    }).then((response) => {
      if (!response.ok) {
        throw new Error("Error sending email: " + response.status);
      }
      return response;
    })
      .then((emailResponse) => {
        console.log("Email sent successfully:", emailResponse);
      })
      .catch((error) => {
        console.error("Error:", error);
      });
  }

  return (
    <>
      {questions.map((question, index) => (
        <Card key={index} style={{ margin: '2%', marginTop: '8px', borderLeft: '3px solid rgb(8 127 231)', backgroundColor: ' #f3f2f2' }}>
          {question}
          <RadioGroup layout="horizontal" onChange={(e) => handleRadioChange(index, e.target.value)}>
            <Radio label='Sí' value='Sí' />
            <Radio label='No' value='No' />
          </RadioGroup>
        </Card>
      ))}

      {showResult && (
        <div style={{ margin: '2%', backgroundColor: 'lemonchiffon', padding: '15px', borderRadius: '8px' }}>
          <p style={{ fontWeight: 'bold' }}>{riskLevel}</p>
          {riskLevel === 'BAJO' && (
            <p>
              Gracias por denunciar el correo como phishing. Según tus respuestas,
              descartamos una situación de riesgo para tu equipo.
              No tienes que hacer nada más. Puedes seguir trabajando con normalidad.
            </p>
          )}
          {riskLevel === 'MEDIO' && (
            <p>
              Gracias por denunciar el correo como phishing.
              Un técnico del CAU contactará contigo para analizar
              tu equipo y descartar que se haya visto infectado.
              Por precaución, cambia la contraseña de acceso a Windows.
            </p>
          )}
          {riskLevel === 'ALTO' && (
            <p>
              Gracias por denunciar el correo como phishing.
              Cambia urgentemente la contraseña de acceso a Windows.
              Un técnico del CAU contactará contigo para analizar tu equipo.
              Estate atento a tu móvil.
            </p>
          )}
          {riskLevel === 'CRITICO' && (
            <p>
              Gracias por denunciar el correo como phishing.
              Activa el modo avión de tu portátil.
              Quita los datos móviles y la wifi de tu móvil.
              Un técnico del CAU te llamará para revisar tu equipo y/o móvil.
            </p>
          )}
        </div>
      )}

      <div id="sendEmailButton" style={{ display: 'flex', justifyContent: 'center', margin: '4%' }}>
        <Button appearance="primary" onClick={handleSendClick} icon={<Send16Filled />} iconPosition="after">
          Enviar
        </Button>
      </div>

    </>
  )
}
