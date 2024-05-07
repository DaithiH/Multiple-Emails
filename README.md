## Python Script To Send Multiple Emails  
 
I had the idea to create this to help my wife send the same email to multiple recipients when she was running a film festival.  
It is necessary to import *smtplib*, the Python Secure Mail Transfer Protocol library, to create a server object and establish a secure connection.
```
server = smtplib.SMTP_SSL("smtp.gmail.com", 465)
    server.ehlo()
    server.login(my_email, my_password)
```
A server object is created, and secure SSL/TLS connection established.
I was using gmail to send the mails, so the connection is through port 465, this will be different for other services. Also, permission will need to be given to allow access for less secure third party apps in the account settings.
Extended Hello, *ehlo()*, method called to identify client, *i.e.* the sender, and request the server capabilities.  
Authenticate the SMTP server with the login() method, and user credentials.   

The list of recipient names and email addresses are kept in an Excel file. The *Pandas* library is imported to read this using the *read_excel()* method. The columns are titled 'Name' and 'Email', but obviously these can be changed.

I created and ran this in the Anaconda Spyder IDE, but it will run using the termial, or whatever Python interpreter or IDE you prefer.




