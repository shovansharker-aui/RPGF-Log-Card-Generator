"""
mail_config.py

Default details for the weekly Maintenance Schedule email draft.

Recipients are left empty on purpose: set To / CC / Subject / Body in the
app (Maintenance Schedule > Settings). They are saved in settings.json.

{date_long} is replaced with the date the mail is prepared, written like
"06th October, 2026".
"""

TO = ""

CC = ""

SUBJECT = "Maintenance Work Schedule"

# Plain paragraphs; each item becomes one paragraph.
BODY = [
    "Dear Concern,",
    "Please find the maintenance schedule given on {date_long}.",
]

# The signature saved in Outlook is inserted automatically. This is used
# only if Outlook has no default signature.
SIGNATURE = []
