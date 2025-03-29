import { SMSParams } from "../models";
import { APIResponse, RequestService } from "../services/request.service";
import { SmsActivityModule } from "./sms/Activity.module";
import { SmsInboundModule } from "./sms/Inbound.module";
import { SmsMessageModule } from "./sms/Message.module";
import { SmsNumberModule } from "./sms/Number.module";
import { SmsRecipientModule } from "./sms/Recipient.module";
import { SmsWebhookModule } from "./sms/Webhook.module";

export class SMSModule extends RequestService  {
  activity: SmsActivityModule;
  number: SmsNumberModule;
  message: SmsMessageModule;
  inbound: SmsInboundModule;
  recipient: SmsRecipientModule;
  webhook: SmsWebhookModule;

  constructor(apiKey: string, baseUrl: string) {
    super(apiKey, baseUrl);

    this.activity = new SmsActivityModule(apiKey, baseUrl);
    this.number = new SmsNumberModule(apiKey, baseUrl);
    this.message = new SmsMessageModule(apiKey, baseUrl);
    this.inbound = new SmsInboundModule(apiKey, baseUrl);
    this.recipient = new SmsRecipientModule(apiKey, baseUrl);
    this.webhook = new SmsWebhookModule(apiKey, baseUrl);
  }

  async send(params: SMSParams): Promise<APIResponse> {
    return await this.post("/sms", params);
  }
}
