#!/usr/bin/python
import time
from exchangelib import DELEGATE, Account, Credentials, Configuration

CHECK_INTERVAL_MINS = 2
USERNAME = "user@domain.tld"
PASSWORD = "password"
SERVER = "outlook.com"

def connect_to_exchange():
    credentials = Credentials(username=USERNAME,
                              password=PASSWORD)
    config = Configuration(server=SERVER, credentials = credentials)
    account = Account(primary_smtp_address=USERNAME,
                      config=config,
                      autodiscover=False,
                      access_type=DELEGATE)
    return account

def set_icon_notify(notify):
    print notify

old_unread = None
notifying = False
quitting = False
account = connect_to_exchange()

while not quitting:
    account.inbox.refresh()
    unread = account.inbox.unread_count

    if old_unread == None:
        notifying = False
    elif notifying:
        notifying = old_unread <= unread
    else:
        notifying = old_unread < unread

    set_icon_notify(notifying)
    old_unread = unread
    time.sleep(CHECK_INTERVAL_MINS * 60)
