EmailNotify
===========

A simple script to send emails from the command line on Windows

Usage
-----

* Syntax: `EmailNotify.exe <subject> <line>`
* Text that contains spaces must be enclosed in quotes
* The message will be sent as raw text, make sure to escape special characters in the message.
* Make sure to set up recipients and server information in config.ini before using.


Config.ini Options
------------------

| Key          | Value                                                                 |
|--------------|-----------------------------------------------------------------------|
|smtpserver:   |[SMTP server, gmail is smtp.gmail.com]                                 |
|username:     |[Your username]                                                        |
|password:     |[Your password]                                                        |
|port:         |[Port to use, for gmail+SSL use 465]                                   |
|ssl:          |[Use SSL to deliver message, 0/1]                                      |
|fromaddress:  |[The address the email will come from]                                 |
|toaddress:    |[Addresses to deliver the email to, separate multiple addresses with ;]|
|ccaddress:    |[Addresses to CC, separate multiple addresses with ;]                  |
|bccaddress:   |[Addresess to BCC, separate multiple addresses with ;]                 |
|importance:   |[Low/Normal/High]                                                      |
|fromname:     |[The name the email will be from]                                      |
