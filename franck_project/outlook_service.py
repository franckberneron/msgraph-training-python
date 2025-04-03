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
from msgraph.generated.users.item.messages.item.create_reply.create_reply_post_request_body import CreateReplyPostRequestBody


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


# Gestion de la boite mail

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

# Opérations d'envoi et de réponse aux mails

## Fonction utilitaires

    def build_mail_html(self, html_body: str, html_signature: str) -> str:
        return f"""
        <html>
          <body style=\"font-family:Segoe UI, sans-serif; font-size:14px; color:#000;\">
            {html_body}
            <br>{html_signature}
          </body>
        </html>
        """

    def _prepare_message(self, html_body: str, inline_images: list = None) -> Message:
        body = ItemBody(
            content_type=BodyType.Html,
            content=html_body
        )

        message = Message(body=body)

        if inline_images:
            attachments = []
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

            message.attachments = attachments

        return message



    async def send_mail(self, sender_user_id: str, to_address: str, subject: str, html_body: str, inline_images: list = None, reply_to_address: tuple = None):
        message = self._prepare_message(html_body, inline_images)

        message.subject = subject
        message.to_recipients = [
            Recipient(email_address=EmailAddress(address=to_address))
        ]

        if reply_to_address:
            email, name = reply_to_address
            message.reply_to = [
                Recipient(email_address=EmailAddress(address=email, name=name))
            ]

        mail_request = SendMailPostRequestBody(
            message=message,
            save_to_sent_items=True
        )

        await self.app_client.users.by_user_id(sender_user_id).send_mail.post(body=mail_request)


    async def reply_to_message(self, user_id: str, message_id: str, html_body: str, inline_images: list = None):
        # Créer le brouillon
        reply_msg = await self.app_client.users.by_user_id(user_id) \
            .messages.by_message_id(message_id) \
            .create_reply.post(
                body=CreateReplyPostRequestBody(comment="")
            )

        # Remplir le contenu et les pièces jointes
        updated_msg = self._prepare_message(html_body, inline_images)

        # Patch du message
        await self.app_client.users.by_user_id(user_id) \
            .messages.by_message_id(reply_msg.id) \
            .patch(body=updated_msg)

        # Envoi
        await self.app_client.users.by_user_id(user_id) \
            .messages.by_message_id(reply_msg.id) \
            .send.post()


    # Réponse à un nombre donnée d'emails
    # Sert à chauffer les adresses mail
     
    async def reply_to_mails(self, user_id: str, nb_messages: int, html_message: str, avg_delay: int = 60, inline_images: list = None):
        messages = await self.get_inbox(user_id, limit=nb_messages)
        cibles = messages[:nb_messages]

        for idx, msg in enumerate(cibles, 1):
            print(f"\n⏳ Message {idx}/{nb_messages} - ID: {msg['id']}")
            print("  De :", msg['from'])
            print("  Objet :", msg['subject'])

            await self.reply_to_message(user_id=user_id, message_id=msg["id"], html_body=html_message, inline_images=inline_images)
            print("✅ Réponse envoyée")

            if idx < nb_messages:
                wait_time = max(5, int(random.normalvariate(avg_delay, 30)))
                print(f"⏱️ Attente de {wait_time} secondes...")
                time.sleep(wait_time)

        print("\n📬 Toutes les réponses ont été envoyées.")


    # Réponse à des mails identifiés
    # Servira à produire des réponses automatiques ciblées

    async def reply_to_email(self, user_id: str, message_ids: list[str], html_message: str, avg_delay: int = 60, inline_images: list = None):
        total = len(message_ids)

        for idx, message_id in enumerate(message_ids, 1):
            print(f"\n⏳ Message {idx}/{total} - ID: {message_id}")

            # Récupération facultative du message pour affichage
            msg = await self.app_client.users.by_user_id(user_id).messages.by_message_id(message_id).get()
            sender = msg.from_.email_address.address if msg and msg.from_ and msg.from_.email_address else "inconnu"
            subject = msg.subject or "(sans objet)"
            print("  De :", sender)
            print("  Objet :", subject)

            # Envoi de la réponse
            await self.reply_to_message(
                user_id=user_id,
                message_id=message_id,
                html_body=html_message,
                inline_images=inline_images
            )

            print("✅ Réponse envoyée")

            # Attente aléatoire avant la prochaine réponse
            if idx < total:
                wait_time = max(5, int(random.normalvariate(avg_delay, 30)))
                print(f"⏱️ Attente de {wait_time} secondes...")
                time.sleep(wait_time)

        print("\n📬 Toutes les réponses ont été envoyées.")