# Gmail to Joplin Using Python

This script retrieves unread e-mails (from, subject, main text, and attachments) from a Gmail account. It then imports them to a notebook Joplin, and initiates Joplin synchronization. **NOTE:** For something more stable that probably does what you need and more right away, see https://github.com/manolitto/joplin-mail-gateway or https://github.com/Hegghammer/joplin-helpers.

In the long run, I aim to re-work this into more tidy code and use my Joplin CLI API. My aim is to enable a few additional commands in e-mails, e.g. set_joplin_notebook="target notebook" for automated sorting to various notebooks. Maybe also some Markdown format options. We'll see.

- 27.01.2022 -- Added call to exit() at the end. 
- 08.01.2022 -- Improved decoding of mail texts, fixing known bug when decoding special characters.
- 06.01.2022 -- Code a bit more up to PEP, and using Pythons native logging. Known bug: encode/decode-trouble sometimes makes mails a bit messy
- 02.01.2022 -- Initial commit. Works fine in tests, but bugs likely. 
