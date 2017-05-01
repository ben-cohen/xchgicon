#!/usr/bin/python
#
# xchgicon.py: Notification icon for unread Exchange emails

import time
import pygtk
import glib
import gtk
import exchangelib
import sys

if len(sys.argv) != 3:
    print "Usage: %s <username> <password>"%sys.argv[0]
    sys.exit(1)

CHECK_INTERVAL_MINS = 2
username = sys.argv[1]
password = sys.argv[2]
server = "outlook.com"

def connect_to_exchange():
    # See https://github.com/ecederstrand/exchangelib for documentation on
    # exchangelib.
    credentials = exchangelib.Credentials(username=username,
                                          password=password)
    config = exchangelib.Configuration(server=server,
                                       credentials = credentials)
    account = exchangelib.Account(primary_smtp_address=username,
                                  config=config,
                                  autodiscover=False,
                                  access_type=exchangelib.DELEGATE)
    return account

def menu_quit_callback(data=None):
    gtk.main_quit()

def show_menu(event_button, event_time, data=None):
    menu = gtk.Menu()
    menu_quit = gtk.MenuItem("Quit")
    menu.append(menu_quit)
    menu_quit.connect_object("activate", menu_quit_callback, "Quit")
    menu_quit.show()
    menu.popup(None, None, None, event_button, event_time)

def on_right_click(data, event_button, event_time):
    show_menu(event_button, event_time)

def set_icon_notify(notify):
    global trayicon
    if notify:
        trayicon.set_from_icon_name("indicator-messages-new")
    else:
        trayicon.set_from_icon_name("indicator-messages")

def perform_check(data):
    global notifying
    global old_unread

    account.inbox.refresh()
    unread = account.inbox.unread_count

    # Say that we need to notify for "new" unread messages if the count goes
    # up, and that we can clear the notification if the count goes down.
    if old_unread == None:
        notifying = False
    elif notifying:
        notifying = old_unread <= unread
    else:
        notifying = old_unread < unread

    #print "DEBUG: %s -> %s so notifying := %s"%(old_unread, unread, notifying)

    set_icon_notify(notifying)
    old_unread = unread

    return True

old_unread = None
notifying = False
account = connect_to_exchange()

trayicon = gtk.status_icon_new_from_icon_name("indicator-messages")
trayicon.connect('popup-menu', on_right_click)

# Perform the check now and at regular intervals thereafter
perform_check(None)
glib.timeout_add_seconds(CHECK_INTERVAL_MINS * 60, perform_check, None)

gtk.main()
sys.exit(0)
