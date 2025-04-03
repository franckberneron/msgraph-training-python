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
# from msgraph.generated.users.item.mail_folders.item.messages import MessagesRequestBuilder

from msgraph.generated.users.item.messages.messages_request_builder import MessagesRequestBuilder
from msgraph.generated.users.item.send_mail.send_mail_post_request_body import SendMailPostRequestBody

class OutlookService:

    settings: SectionProxy
    client_credential: ClientSecretCredential
    app_client: GraphServiceClient

    def __init__(self, config):
        # Initialise le client Graph avec un token d'accès
        self.settings = config
        client_id = self.settings['clientId']
        tenant_id = self.settings['tenantId']
        client_secret = self.settings['clientSecret']

        self.client_credential = ClientSecretCredential(tenant_id, client_id, client_secret)
        self.app_client = GraphServiceClient(self.client_credential)  # type: ignore

    async def get_inbox(self, user_id: str):
        """
        Récupère les 25 derniers emails avec id, expéditeur, objet et statut lu/non-lu.
        :param user_id: ID ou email de l'utilisateur
        :return: Liste de dictionnaires
        """
        query_params = MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
            top=25,
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
            """
            Envoie un email avec corps HTML, images inline et adresse de réponse personnalisée.

            :param sender_user_id: L'utilisateur expéditeur (email ou ID Azure).
            :param to_address: Adresse email du destinataire.
            :param subject: Sujet du mail.
            :param html_body: Corps HTML du mail (avec balises cid:).
            :param inline_images: Liste de tuples (chemin_image, content_id).
            :param reply_to_address: Tuple (adresse email, nom affiché) pour réponse.
            """
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
                    attachment.name = image_path.split("/")[-1]  # nom du fichier
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
        """
        Répond à un message existant avec une réponse formatée en HTML.

        :param user_id: Identifiant ou email de l'utilisateur.
        :param message_id: ID du message auquel répondre.
        :param html_body: Corps HTML de la réponse.
        """
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
