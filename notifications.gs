/*Éste código automatiza el envío de correos electrónicos y sms al usuario después de llenar el formulario (en google form).

Para implementarlo es necesario en el spreedsheet de google form agregar una secuencia de comandos
(Herramientas -> Editor de secuencia de comandos) dónde se inserta el código.

Posteriormente desde esa misma ventana hay que ir a la parte de Activadores (Editar -> Activadores de proyecto activo),
dónde se agrega un nuevo activador para la función onFormSubmit cuando se envíe el formulario.

Descripción general del código:

- sendSms: Función que implementa el API de Twilio para enviar un sms.
- sendEmailGoogle (deprecado): Función que implementa el API nativa de gmail para enviar correos electrónicos
  (requiere permisos de la cuenta del autor para envío de correos electrónicos).
- sendEmailSendgrid: Función que implementa el API de Sendgrid para enviar correos electrónicos.
- onFormSubmit: Función que atrapa el registro insertado y envía una notificación de acuerdo a las reglas de clasificación de
  riesgo.

Los correos electrónicos adjuntan un PDF el cual es descargado de google drive.
*/

var LOW_TEMPLATE_ID = "<GOOGLE-DRIVE-PDF-ID>"; 
var INTERMEDIATE_TEMPLATE_ID = "<GOOGLE-DRIVE-PDF-ID>"; 
var HIGH_TEMPLATE_ID = "<GOOGLE-DRIVE-PDF-ID>"; 

function sendSms(recipient, name, risk_type) {
  var messages_url = "https://api.twilio.com/2010-04-01/Accounts/<TWILIO-ID>/Messages.json";
  var risk = '';
  var url = '';
  
  switch (risk_type) {
    case "L":
      risk = 'BAJO';
      url = 'https://drive.google.com/file/d/' + LOW_TEMPLATE_ID;
      break;
    case "I":
      risk = 'INTERMEDIO';
      url = 'https://drive.google.com/file/d/' + INTERMEDIATE_TEMPLATE_ID;
      break;
    case "H":
      risk = 'ALTO';
      url = 'https://drive.google.com/file/d/' + HIGH_TEMPLATE_ID;
      break;
  }
  
  var body = 'Hola '+ name +', tu riesgo de ser un caso de COVID-19 es ' + risk + ', te compartimos las siguientes recomendaciones: ' + url + '. Gracias por compartir la encuesta, Saludos.';

  var payload = {
    "To": "+52" + recipient,
    "Body" : body,
    "From" : "<TWILIO-CELLPHONE>"
  };

  var options = {
    "method" : "post",
    "payload" : payload
  };

  options.headers = { 
    "Authorization" : "Basic " + Utilities.base64Encode("<TWILIO-TOKEN>:<TWILIO-TOKEN>")
  };

  var response = UrlFetchApp.fetch(messages_url, options);
  return response;
}

function sendEmailGoogle(recipient, name, risk_type) {
  var EMAIL_SUBJECT = 'COVID-19 GDL';
  var email_body = '';
  var template_id = '';
  var file_name = '';
  
  switch (risk_type) {
    case "L":
      template_id = LOW_TEMPLATE_ID;
      email_body = '<P>Hola <b><span style="color: red;">'+ name +'</span></b>, tu riesgo de ser un caso de COVID-19 es <b><span style="color: red;">BAJO</span></br>, te compartimos las siguientes recomendaciones.</P></BR><P>Gracias por compartir la encuesta, Saludos.</P>';
      file_name = 'Riesgo_Bajo.pdf';
      break;
    case "I":
      template_id = INTERMEDIATE_TEMPLATE_ID;
      email_body = '<P>Hola <b><span style="color: red;">'+ name +'</span></b>, tu riesgo de ser un caso de COVID-19 es <b><span style="color: red;">INTERMEDIO</span></br>, te compartimos las siguientes recomendaciones.</P></BR><P>Gracias por compartir la encuesta, Saludos.</P>';
      file_name = 'Riesgo_Intermedio.pdf';
      break;
    case "H":
      template_id = HIGH_TEMPLATE_ID;
      email_body = '<P>Hola <b><span style="color: red;">'+ name +'</span></b>, tu riesgo de ser un caso de COVID-19 es <b><span style="color: red;">ALTO</span></br>, te compartimos las siguientes recomendaciones.</P></BR><P>Gracias por compartir la encuesta, Saludos.</P>';
      file_name = 'Riesgo_Alto.pdf';
      break;
  }
  
  email_body = '<HTML><BODY>' + email_body + '</BODY></HTML>';
  
  
  var copyFile = DriveApp.getFileById(template_id).makeCopy();
  var copyId = copyFile.getId();
  
  var newFile = DriveApp.createFile(copyFile.getAs('application/pdf'));
  newFile.setName(file_name);
  copyFile.setTrashed(true);
  MailApp.sendEmail({
    to:recipient,
    subject:EMAIL_SUBJECT,
    htmlBody:email_body,
    attachments: [newFile]});
}

function sendEmailSendgrid(recipient, name, risk_type) {
  var SENDGRID_KEY ='<SENDGRID-KEY>';
  var EMAIL_SUBJECT = 'COVID-19 GDL';
  var EMAIL_FROM = 'covid19gdl@covid19gdl.com';
  var template_id = '';
  var file_name = '';
  var risk = '';
  
  switch (risk_type) {
    case "L":
      template_id = LOW_TEMPLATE_ID;
      file_name = 'Riesgo_Bajo.pdf';
      risk = 'BAJO';
      break;
    case "I":
      template_id = INTERMEDIATE_TEMPLATE_ID;
      file_name = 'Riesgo_Intermedio.pdf';
      risk = 'INTERMEDIO';
      break;
    case "H":
      template_id = HIGH_TEMPLATE_ID;
      file_name = 'Riesgo_Alto.pdf';
      risk = 'ALTO';
      break;
  }
  
  var email_body = '<P>Hola <span style="color: red; font-weight: bold;">'+ name +'</span>, tu riesgo de ser un caso de COVID-19 es <span style="color: red; font-weight: bold;">' + risk + '</span>, te compartimos las siguientes recomendaciones.</P></BR><P>Gracias por compartir la encuesta, Saludos.</P>';
  
  email_body = '<HTML><BODY>' + email_body + '</BODY></HTML>';
  var file = DriveApp.getFileById(template_id);
  var base64EncodedBytes = Utilities.base64Encode(file.getBlob().getBytes()); 

  var headers = {
    "Authorization" : "Bearer " + SENDGRID_KEY, 
    "Content-Type": "application/json" 
  }
  
  var body =
  {
    "personalizations": [
      {
        "to": [
          {
            "email": recipient
          }
        ],
        "subject": EMAIL_SUBJECT
      }
    ],
    "from": {
      "email": EMAIL_FROM
    },
    "content": [
      {
        "type": "text/html",
        "value": email_body
      }
    ],
    "attachments": [
      {
        "content": base64EncodedBytes,
        "type": "application/pdf",
        "filename": file_name,
        "disposition": "attachment"
      }
    ]
  }
  
  
  var options = {
    'method': 'post',
    'headers': headers,
    'payload': JSON.stringify(body)
  }

  var response = UrlFetchApp.fetch("https://api.sendgrid.com/v3/mail/send", options);

  Logger.log(response);
  return response;
}

function onFormSubmit(e) {
  var countries_in_risk = ['España', 'Alemania', 'Francia', 'Italia', 'Estados Unidos', 'China', 'Japón', 'Inglaterra'];
  var sheet = SpreadsheetApp.getActiveSheet();
  var name = e.values[3].toUpperCase();
  var telephone = e.values[7];
  var email = e.values[1];
  var score = parseInt(e.values[2].slice(0, 2));
  var countries_traveled = e.values[28].split(",");
  var is_country_in_risk = false;
  var coronavirus_friend = e.values[25];
  var is_excited = e.values[19];
  var can_finished_phrases = e.values[20];
  
  try {
    for (country in countries_traveled) {
      if (countries_traveled.includes(country.trim())) {
        is_country_in_risk = true;
        break;
      }
    }
    
    if ((score >= 4 && score <= 6) || is_country_in_risk || coronavirus_friend === 'Sí' || (is_excited === 'Sí' && can_finished_phrases === 'Sí')) {
      risk_type = 'I';
      color = 'yellow';
    } else if (score >= 7 || (is_excited === 'Sí' && can_finished_phrases === 'No')) {
      risk_type = 'H';
      color = 'red';
    } else if (score <= 4) {
      risk_type = 'L';
      color = '#00ff00';
    }
   
    response_data_email = sendEmailSendgrid(email, name, risk_type);
    //response_data_sms = sendSms(cellphone, name, risk_type);
    status = "Enviado Automático"
  } catch(err) {
    Logger.log(err);
    status = "Error en el envío automático";
  }
  sheet.getRange(sheet.getLastRow(), 31).setValue(status);
  sheet.getRange(sheet.getLastRow(), 1, 1, 31).setBackground(color);
}
