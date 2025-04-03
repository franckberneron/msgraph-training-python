from configparser import SectionProxy
from msgraph import GraphServiceClient
from azure.identity.aio import ClientSecretCredential

from msgraph.generated.models.body_type import BodyType
from msgraph.generated.models.message import Message
from msgraph.generated.models.email_address import EmailAddress
from msgraph.generated.models.importance import Importance
from msgraph.generated.models.item_body import ItemBody
from msgraph.generated.models.recipient import Recipient
from msgraph.generated.models.file_attachment import FileAttachment
from msgraph.generated.users.item.messages.messages_request_builder import MessagesRequestBuilder
from msgraph.generated.users.item.send_mail.send_mail_post_request_body import SendMailPostRequestBody

import time
import random

class OutlookService:

    settings: SectionProxy
    client_credential: ClientSecretCredential
    app_client: GraphServiceClient

    def __init__(self, config):
        self.settings = config
        client_id = self.settings['clientId']
        tenant_id = self.settings['tenantId']
        client_secret = self.settings['clientSecret']

        self.client_credential = ClientSecretCredential(tenant_id, client_id, client_secret)
        self.app_client = GraphServiceClient(self.client_credential)  # type: ignore

    async def get_inbox(self, user_id: str, limit: int = 25):
        query_params = MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
            top=limit,
            orderby=["receivedDateTime DESC"],
            select=["id", "from", "subject", "isRead"]
        )

        request_config = MessagesRequestBuilder.MessagesRequestBuilderGetRequestConfiguration(
            query_parameters=query_params
        )

        messages = await self.app_client.users.by_user_id(user_id) \
            .mail_folders.by_mail_folder_id("inbox") \
            .messages.get(request_configuration=request_config)

        result = []
        for msg in getattr(messages, 'value', []):
            result.append({
                "id": msg.id,
                "from": getattr(msg.from_, 'email_address', {}).address if msg.from_ and msg.from_.email_address else None,
                "subject": msg.subject,
                "isRead": msg.is_read
            })

        return result

    async def send_mail(self, sender_user_id: str, to_address: str, subject: str, html_body: str, inline_images: list = None, reply_to_address: tuple = None):
        body = ItemBody(
            content_type=BodyType.Html,
            content=html_body
        )

        recipients = [
            Recipient(
                email_address=EmailAddress(address=to_address)
            )
        ]

        reply_to = []
        if reply_to_address:
            email, name = reply_to_address
            reply_to = [
                Recipient(
                    email_address=EmailAddress(
                        address=email,
                        name=name
                    )
                )
            ]

        message = Message(
            subject=subject,
            body=body,
            to_recipients=recipients,
            reply_to=reply_to
        )

        attachments = []
        if inline_images:
            for image_path, content_id in inline_images:
                with open(image_path, "rb") as f:
                    raw_image = f.read()

                attachment = FileAttachment()
                attachment.name = image_path.split("/")[-1]
                attachment.content_type = "image/png"
                attachment.is_inline = True
                attachment.content_id = content_id
                attachment.content_bytes = raw_image

                attachments.append(attachment)

        if attachments:
            message.attachments = attachments

        mail_request = SendMailPostRequestBody(
            message=message,
            save_to_sent_items=True
        )

        await self.app_client.users.by_user_id(sender_user_id).send_mail.post(body=mail_request)

    async def reply_to_message(self, user_id: str, message_id: str, html_body: str):
        body = ItemBody(
            content_type=BodyType.HTML,
            content=html_body
        )

        await self.app_client.users.by_user_id(user_id)\
            .messages.by_message_id(message_id)\
            .reply.post(body={"comment": html_body})

        await self.app_client.users.by_user_id(user_id)\
            .messages.by_message_id(message_id)\
            .send.post()

    def construire_mail_html(self) -> str:
        signature_html = """
        <p><strong>Franck BERNERON</strong><br>
        Associé co-fondateur</p>
        <p><strong>Mobile</strong> +33 756 26 73 03<br>
        <strong>Email</strong> <a href=\"mailto:franck@axires.tech\">franck@axires.tech</a><br>
        <strong>Adresse</strong> 96, rue Paradis<br>
        13006 Marseille</p>
        <img src=\"cid:logo\" width=\"120\">
        """

        return f"""
        <html>
          <body style=\"font-family:Segoe UI, sans-serif; font-size:14px; color:#000;\">
            <p>Bonjour,</p>
            <p>Merci pour votre message. Nous vous répondrons dans les meilleurs délais.</p>
            <p>Très cordialement,</p>
            <br>{signature_html}
          </body>
        </html>
        """

    async def repondre_mails(self, user_id: str, nb_messages: int, delai_moyen: int = 60):
        messages = await self.get_inbox(user_id)
        cibles = messages[:nb_messages]

        for idx, msg in enumerate(cibles, 1):
            print(f"\n⏳ Message {idx}/{nb_messages} - ID: {msg['id']}")
            print("  De :", msg['from'])
            print("  Objet :", msg['subject'])

            html = self.construire_mail_html()
            await self.reply_to_message(user_id=user_id, message_id=msg["id"], html_body=html)
            print("✅ Réponse envoyée")

            if idx < nb_messages:
                attente = max(5, int(random.normalvariate(delai_moyen, 30)))
                print(f"⏱️ Attente de {attente} secondes...")
                time.sleep(attente)

        print("\n📬 Toutes les réponses ont été envoyées.")
