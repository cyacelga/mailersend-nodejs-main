const { MailerSend } = require("mailersend");
const json = require("./example.json");

const mailersend = new MailerSend({
  apiKey: "TOKEN_API_MAILSENDER",
});

const emailParams = {
  from: {
    email: json.from,
    name: json.fromName
  },
  to: [
    {
      email: json.to,
      name: json.name
    }
  ],
  subject: json.subject,
  template_id: json.templateId
};

mailersend.email.send(emailParams)
  .then(() => console.log("📨 Correo enviado correctamente"))
  .catch((err) => console.error("❌ Error al enviar el correo:", err));
