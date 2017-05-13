import smtplib
from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText

fromaddr = "tien.nguyen@viethope.org"
toaddr = "thu.nguyen@viethope.org"
cc = "tien.nguyen@viethope.org"

msg = MIMEMultipart()
msg['From'] = fromaddr
msg['To'] = toaddr
msg['Subject'] = "Thank you for your support!"

body = "Testing 1234!"
msg.attach(MIMEText(body, 'plain'))

server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.login(fromaddr, "")
text = msg.as_string()

for i in xrange(0, 10):
    server.sendmail(fromaddr, [toaddr, cc], text)

server.quit()
