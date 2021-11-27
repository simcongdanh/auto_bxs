import imaplib
import email, re

# Connect to inbox
imap_server = imaplib.IMAP4_SSL(host='outlook.office365.com')
imap_server.login('amestjbalsermdl@outlook.com', '9S2aTJ87')
#print(imap_server.list())
imap_server.select("Junk")  # Default is `INBOX`

# Find all emails in inbox
_, message_numbers_raw = imap_server.search(None, 'ALL')
for message_number in message_numbers_raw[0].split():
    _, msg = imap_server.fetch(message_number, '(RFC822)')

    # Parse the raw email message in to a convenient object
    message = email.message_from_bytes(msg[0][1])
    print('== Email message =====')
    # print(message)  # print FULL message
    print('== Email details =====')
    # print(type(message["Content-Type"]))
    # link = re.match(r"(<http\w+>)", message["Content-Type"])
    # print(f'From: {message["from"]}')
    # print(f'To: {message["to"]}')
    # print(f'Cc: {message["Content-Type"]}')
    # print(f'Bcc: {message["subject"]}')
    # # get_link = re.compile("<http(.*)>")
    # # result = get_link.search(message)
    # #print(link)
    # print(f'Urgency (1 highest 5 lowest): {message["x-priority"]}')
    # print(f'Object type: {type(message)}')
    # print(f'Content type: {message.get_content_type()}')
    # print(f'Content disposition: {message.get_content_disposition()}')
    # print(f'Multipart?: {message.is_multipart()}')
    if message.is_multipart():
        print('Multipart types:')
        for part in message.walk():
            print(f'- {part.get_content_type()}')
        multipart_payload = message.get_payload()
        for sub_message in multipart_payload:
            # The actual text/HTML email contents, or attachment data
            # print(f'Payload\n{sub_message.get_payload()}')
            # print(type(sub_message.get_payload()))
            lines = sub_message.get_payload().split('\n')
            for i in range(0, len(lines)):
                for line in lines:
                    print(f"line-{i}-{line[i]}")
    else:  # Not a multipart message, payload is simple string
        # break
        print(f'Payload\n{message.get_payload()}')
        print(type(message.get_payload()))
        lines = message.get_payload().split('\n')
        for i in range(0, len(lines)):
            for line in lines:
                print(f"line-{i}-{line[i]}")
    # You could also use `message.iter_attachments()` to get attachments only
