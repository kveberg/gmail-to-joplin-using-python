# joplin-helpers
Helpers for Joplin (see https://joplinapp.org/).

## Gmail to Joplin (Python)
This script retrieves unread e-mails (from, subject, main text, and attachments) from a Gmail account. It then imports them to a notebook Joplin, and initiates Joplin synchronization. **NOTE:** For something more stable that probably does what you need and more right away, see https://github.com/manolitto/joplin-mail-gateway or https://github.com/Hegghammer/joplin-helpers.

- 06.01.2022 -- Code a bit more up to PEP, and using Pythons native logging. Known bug: encode/decode-trouble sometimes makes mails a bit messy
- 02.01.2022 -- Initial commit. Works fine in tests, but bugs likely. 
