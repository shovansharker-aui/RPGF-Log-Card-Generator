"""
mail_config.py

Default details for the weekly Maintenance Schedule email draft.

These are PLACEHOLDERS. Real recipients are set in the app
(Maintenance Schedule > Settings) or in a private mail_config_local.py,
which is git-ignored and overrides everything below.

{date_long} is replaced with the date the mail is prepared, written like
"06th October, 2026".
"""

TO = "recipient@example.com"

CC = "cc.person@example.com"

SUBJECT = "Maintenance Work Schedule"

# Plain paragraphs; each item becomes one paragraph.
BODY = [
    "Dear Concern,",
    "Please find the maintenance schedule given on {date_long}.",
]

# The signature saved in Outlook is inserted automatically. This is used
# only if Outlook has no default signature.
SIGNATURE = []

try:
    from mail_config_local import *  # noqa: F401,F403
except ImportError:
    pass
