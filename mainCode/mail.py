import logging
import smtplib
import email.mime.multipart
import email.mime.text
import email.mime.base
import email.encoders as encoders



def send_mail(receiver_email: str, mail_subject: str, mail_body: str, attachment_locations: list = None, sender_email: str = None, sender_password: str=None) -> bool:
    """The Function responsible to do all the mail sending logic.

    Args:
        sender_email (str): Email Id of the sender.
        sender_password (str): Password of the sender.
        receiver_email (str): Email Id of the receiver.
        mail_subject (str): Subject line of the email.
        mail_body (str): Message body of the Email.
        attachment_locations (list, optional): Absolute path of the attachment. Defaults to None.

    Returns:
        bool: [description]
    """
    logging.info("INTO THE SEND MAIL FUNCTION")
    done = False
    try:
        logging.info("GIVING CREDENTIALS FOR SENDING MAIL")
        if not sender_email or sender_password:
            sender_email = "biourjapowerdata@biourja.com"
            sender_password = r"Texas08642"
            # sender_email = r"virtual-out@biourja.com"
            # sender_password = "t?%;`p39&Pv[L<6Y^cz$z2bn"
        receivers = receiver_email.split(",")
        msg = email.mime.multipart.MIMEMultipart()
        msg['From'] = "biourjapowerdata@biourja.com"
        msg['To'] = receiver_email
        msg['Subject'] = mail_subject
        body = mail_body
        logging.info("Attaching mail body")
        msg.attach(email.mime.text.MIMEText(body, 'html'))
        logging.info("Attching files in the mail")
        for files_locations in attachment_locations:
            with open(files_locations, 'r+b') as attachment:
                # instance of MIMEBase and named as p
                p = email.mime.base.MIMEBase('application', 'octet-stream')
                # To change the payload into encoded form
                p.set_payload((attachment).read())
                encoders.encode_base64(p)  # encode into base64
                p.add_header('Content-Disposition',
                             "attachment; filename= %s" % files_locations)
                msg.attach(p)  # attach the instance 'p' to instance 'msg'

        # s = smtplib.SMTP('smtp.gmail.com', 587) # creates SMTP session
        s = smtplib.SMTP('smtp.office365.com',
                         587)  # creates SMTP session
        s.starttls()  # start TLS for security
        s.login(sender_email, sender_password)  # Authentication
        text = msg.as_string()  # Converts the Multipart msg into a string

        s.sendmail(sender_email, receivers, text)  # sending the mail
        s.quit()  # terminating the session
        done = True
        logging.info("Email sent successfully")
        print("Email sent successfully.")
    except Exception as e:
        print(
            f"Could not send the email, error occured, More Details : {e}")
    finally:
        return done
