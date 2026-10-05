import os
import re
import uuid


def get_outlook_application():
    """
    Connect to Microsoft Outlook Desktop.

    Returns
    -------
    Outlook.Application
        Outlook COM object used to create and send emails.
    """

    import win32com.client as win32

    return win32.Dispatch("Outlook.Application")


def send_email(
    outlook_app,
    to_address,
    cc_address,
    subject,
    html_body,
    attachment_paths=None
):
    """
    Send one personalized HTML email through Outlook.

    This function also embeds images that came from the Word
    document directly inside the email.

    Flask preview image example:

        /static/uploads/word_images/image123.png

    Before sending, this becomes:

        cid:some_unique_id

    The actual image is then attached invisibly to the email.

    Parameters
    ----------
    outlook_app:
        Outlook application returned by get_outlook_application().

    to_address:
        Recipient email address.

    cc_address:
        Optional CC email address.

    subject:
        Personalized email subject.

    html_body:
        Personalized HTML email body.

    attachment_paths:
        Optional normal attachments.

    Returns
    -------
    tuple:
        (True, None) if successful.

        (False, error_message) if sending fails.
    """

    try:

        # =====================================================
        # Create Outlook email
        # =====================================================

        mail = outlook_app.CreateItem(0)

        mail.To = to_address

        if cc_address:
            mail.CC = cc_address

        mail.Subject = subject


        # =====================================================
        # Copy the HTML so we can modify image paths
        # =====================================================

        final_html = str(html_body)


        # =====================================================
        # Find Word images inside the generated HTML
        # =====================================================
        #
        # Example:
        #
        # <img src="/static/uploads/word_images/abc.png">
        #
        # We only need the filename:
        #
        # abc.png
        # =====================================================

        image_pattern = (
            r'src=["\']'
            r'/static/uploads/word_images/'
            r'([^"\']+)'
            r'["\']'
        )

        image_filenames = re.findall(
            image_pattern,
            final_html,
            flags=re.IGNORECASE
        )


        # =====================================================
        # Embed every image into Outlook
        # =====================================================

        for image_filename in set(image_filenames):

            # Absolute path to the extracted image
            image_path = os.path.join(
                os.getcwd(),
                "static",
                "uploads",
                "word_images",
                image_filename
            )

            # If the file does not exist, skip it.
            if not os.path.exists(image_path):
                continue


            # -------------------------------------------------
            # Create unique Content-ID
            # -------------------------------------------------

            content_id = (
                "oneworld_"
                + uuid.uuid4().hex
            )


            # -------------------------------------------------
            # Attach image to Outlook
            # -------------------------------------------------

            inline_attachment = mail.Attachments.Add(
                os.path.abspath(image_path)
            )


            # -------------------------------------------------
            # Tell Outlook this attachment is an inline image
            # -------------------------------------------------

            property_accessor = (
                inline_attachment.PropertyAccessor
            )


            # PR_ATTACH_CONTENT_ID
            property_accessor.SetProperty(
                "http://schemas.microsoft.com/mapi/"
                "proptag/0x3712001F",
                content_id
            )


            # PR_ATTACHMENT_HIDDEN
            #
            # Makes the image act like part of the email
            # instead of appearing as a normal attachment.
            property_accessor.SetProperty(
                "http://schemas.microsoft.com/mapi/"
                "proptag/0x7FFE000B",
                True
            )


            # =================================================
            # Replace Flask URL with Outlook CID URL
            # =================================================
            #
            # Before:
            #
            # /static/uploads/word_images/abc.png
            #
            # After:
            #
            # cid:oneworld_xxxxx
            # =================================================

            flask_image_url = (
                "/static/uploads/word_images/"
                + image_filename
            )

            cid_image_url = (
                "cid:"
                + content_id
            )

            final_html = final_html.replace(
                flask_image_url,
                cid_image_url
            )


        # =====================================================
        # Set HTML email body
        # =====================================================

        mail.HTMLBody = final_html


        # =====================================================
        # Add normal attachments
        # =====================================================

        for path in (
            attachment_paths or []
        ):

            if os.path.exists(path):

                mail.Attachments.Add(
                    os.path.abspath(path)
                )


        # =====================================================
        # Send email
        # =====================================================

        mail.Send()

        return True, None


    except Exception as exc:

        return False, str(exc)