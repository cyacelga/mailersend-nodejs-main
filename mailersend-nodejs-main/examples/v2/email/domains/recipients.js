import 'dotenv/config';
import { MailerSend } from "mailersend";

const mailerSend = new MailerSend({
  apiKey: process.env.API_KEY,
});

mailerSend.email.domain.recipients("domain_id", {
  page: 1,
  limit: 10
})
  .then((response) => console.log(response.body))
  .catch((error) => console.log(error.body));
