import os
import sys
import json
import smtplib

from optparse import OptionParser, OptionGroup
from datetime import datetime
from hashlib import md5

import gflags
import httplib2

from apiclient.discovery import build
from oauth2client.file import Storage
from oauth2client.client import OAuth2WebServerFlow
from oauth2client.tools import run

CLIENT_ID = "308912144163.apps.googleusercontent.com"
CLIENT_SECRET = "3rcSPCv_sUSC7nPLBDJg8gOy"
USER_AGENT = "CalendarMailer/0.1"

def get_credentials(config, authorize = False):
	flow = OAuth2WebServerFlow(CLIENT_ID, CLIENT_SECRET, scope = "https://www.googleapis.com/auth/calendar.readonly", user_agent = USER_AGENT)
	storage = Storage(os.path.join(config, "oauth.json"))
	if authorize:
		credentials = run(flow, storage)
	else:
		credentials = storage.get()
		if not credentials or credentials.invalid:
			raise Exception("Credentials not found or expired. Please authorize first.")
	return credentials

def get_service(credentials):
	http = httplib2.Http()
	http = credentials.authorize(http)
	return build("calendar", "v3", http = http)

def get_all_calendars(service):
	all = []
	list = service.calendarList().list().execute()
	while True:
		all.extend(list["items"])
		if "nextPageToken" not in list: break
		list = service.calendarList().list(pageToken = list["nextPageToken"]).execute()
	return all

def get_subscriptions(config):
	path = os.path.join(config, "subscriptions.json")
	if not os.path.exists(path):
		return {}
	with open(path, "r") as fp:
		return json.load(fp)

def set_subscriptions(config, subscriptions):
	path = os.path.join(config, "subscriptions.json")
	with open(path, "w") as fp:
		json.dump(subscriptions, fp)

def find_subscription(subscriptions, summary_or_id):
	if summary_or_id in subscriptions:
		return summary_or_id, subscriptions[summary_or_id]
	else:
		for id, subscription in subscriptions.iteritems():
			if subscription["summary"].lower() == summary_or_id.lower():
				return id, subscription
		else:
			print >>sys.stderr, "Error: No calendar with Title or ID '%s' found" % summary_or_id
			return None, None

def get_events(service, calendarId, subscription):
	all = []

	args = {"maxResults": 2, "singleEvents": True, "orderBy": "startTime"}
	for key in ("maxResults", "orderBy", "showDeleted"):
		if key in subscription:
			args[key] = subscription["args"]

	args["timeMin"] = datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%S") + "Z"
	args["fields"]  = "items(id,start,end,summary,description,location,status,updated)"

	list = service.events().list(calendarId = calendarId, **args).execute()
	while True:
		all.extend(list["items"])
		if "nextPageToken" not in list: break
		list = service.events().list(calendarId = calendarId, pageToken = list["nextPageToken"]).execute()
	return all

def get_template(config, subscription):
	def fix_newlines(string):
		return "\r\n".join(string.splitlines())

	if "template" in subscription:
		file = os.path.join(config, subscription["template"])
		if os.path.exists(file):
			with open(file, "r") as fp:
				template = fp.read()
			return fix_newlines(template).split("\r\n\r\n")
		else:
			print >>sys.stderr, "Error: Template '%s' does not exist" % file

	headers = "From: %(from)s\nTo: %(to)s\nSubject: %(subject)s\nContent-Type: text/html"
	body = \
"""
<h1>%(summary)s</h1>
<h2>When</h2>
%(when)s
<h2>Description</h2>
%(description)s
"""
	return fix_newlines(headers), fix_newlines(body)

def send_email(config, subscription, events, dryrun = False):
	def strptime(value):
		if "+" in value: value, tz = value.split("+")
		if value.endswith("Z"): value = value[:-1]
		if "." in value: value, ms = value.split(".")
		return datetime.strptime(value, "%Y-%m-%dT%H:%M:%S")

	if not subscription["recipients"]:
		return

	savedLocale = None
	if "locale" in subscription:
		import locale
		savedLocale = locale.setlocale(locale.LC_TIME, str(subscription["locale"]))

	try:
		dateFormat = subscription["dateFormat"] if "dateFormat" in subscription else "%c"
		vars = {}

		recipients = ["%s <%s>" % (n, e) for e, n in subscription["recipients"].iteritems()]
		vars["to"] = ", ".join(recipients)

		if "from" in subscription:
			vars["from"] = subscription["from"]
		else:
			if os.path.exists("/etc/mailname"):
				with open("/etc/mailname") as fp:
					mailname = fp.read().strip()
			else:
				mailname = os.uname()[1]
			vars["from"] = "Calendar <calmailer@%s>" % mailname

		from_, to = vars["from"], vars["to"]

		if "subject" in subscription:
			vars["subject"] = subscription["subject"]
		else:
			vars["subject"] = subscription["summary"]

		headerTemplate, bodyTemplate = get_template(config, subscription)
		message = headerTemplate.strip() % vars
		message += "\r\n\r\n"

		for event in events:
			updated = strptime(event["updated"])
			start = strptime(event["start"]["dateTime"])
			end = strptime(event["end"]["dateTime"]) if "end" in event else None

			vars["start"]    = start.strftime(dateFormat)
			vars["updated"]  = updated.strftime(dateFormat)
			if end:
				vars["end"]  = end.strftime(dateFormat)
				vars["when"] = "%s - %s" % (vars["start"], vars["end"])
			else:
				vars["end"]  = ""
				vars["when"] = vars["start"]

			for name in ("status", "description", "summary", "location"):
				vars[name] = event[name] if name in event else "-"

			message += bodyTemplate.strip() % vars
			message += "\r\n\r\n"

		hash = md5(subscription["summary"]).hexdigest()
		messageFile = os.path.join(config, hash + ".eml")
		if os.path.exists(messageFile):
			with open(messageFile) as fp:
				oldMessage = fp.read()
			if oldMessage == message:
				if dryrun: print >>sys.stderr, "New message is same as previous, not sending"
				return

		if dryrun:
			print >>sys.stderr, "From:", from_
			print >>sys.stderr, "To:", from_
			print >>sys.stderr, message
		else:
			server = smtplib.SMTP("localhost")
			server.sendmail(from_, to, message)
			server.quit()

			with open(messageFile, "w") as fp:
				fp.write(message)

	finally:
		if savedLocale:
			locale.setlocale(locale.LC_TIME, savedLocale)

def main():

	parser = OptionParser(usage = "usage: %prog [options] command [command options]...")
	parser.add_option("--commands", action = "store_true",
		help = "display available subcommands")
	parser.add_option("-c", "--config", default = "~/.calmailer", metavar = "DIR",
		help = "directory to store configuration files")
	parser.add_option("--dry-run", dest = "dryrun", action = "store_true",
		help = "do everything except sending the actual email message")

	auth = OptionGroup(parser, "OAuth Options")
	auth.add_option("--auth_local_webserver", dest = "auth_local_webserver",
		action = "store_true", default = True, help = "use a local webserver")
	auth.add_option("--noauth_local_webserver", dest = "auth_local_webserver",
		action = "store_false", help = "do not use a local webserver")
	parser.add_option_group(auth)

	options, args = parser.parse_args()

	if options.commands:
		parser.print_help()
		print
		print "Available sub-commands:"
		print "  authorize"
		print "    Initiate OAuth Authorization"
		print
		print "  list"
		print "    List all calendars we are authorized for"
		print
		print "  subscribe CALENDAR"
		print "    Subscribe to a calendar with id or title CALENDAR"
		print
		print "  unsubscribe CALENDAR"
		print "    Un-subscribe from a calendar with id or title CALENDAR"
		print
		print "  subscriptions"
		print "    List all subscriptions"
		print
		print "  add CALENDAR EMAIL NAME"
		print "    Add a recipient with email address EMAIL and name NAME to the subscription"
		print "    for the calendar with id or title CALENDAR"
		print
		print "  remove CALENDAR RECIPIENT"
		print "    Remove a recipient with email address or name RECIPIENT from the "
		print "    subscription for the calendar with id or title CALENDAR"
		print
		print "  set CALENDAR NAME VALUE"
		print "    Set the configuration option NAME to VALUE for the subscription with id or"
		print "    title CALENDAR. NAME can be any of:"
		print "    maxResults"
		print "      Number of events to include in email message (integer, default: 2)"
		print "    showDeleted"
		print "      Include deleted events in email message (true or false, default: false)"
		print "    template"
		print "      The template file to use for the email message (filename relative to "
		print "      configuration directory)"
		print "    subject"
		print "      The Subject for the email message (unless overridden in template)"
		print "    from"
		print "      The Sender for the email message (unless overridden in template)"
		print "    dateFormat"
		print "      strftime format to use for dates (default: %c)"
		print "      (See also http://docs.python.org/library/time.html#time.strftime)"
		print "    locale"
		print "      The locale to use for date and time formatting"
		print
		print "  send CALENDAR"
		print "    Generate email message for the subscription with id or title CALENDAR"
		print
		print "  sendall"
		print "    Generate email messages for all subscriptions"
		print
		print "Note: Multiple sub-commands can be chained. Assuming a calendar named 'Test'"
		print "  exists, the following would be a valid way to subscribe, add a recipient,"
		print "  and send the resulting email message: "
		print
		print "  %s subscribe Test  add Test bob@example.com Bob  send Test" % sys.argv[0]
		print

		sys.exit(0)

	config = os.path.expanduser(options.config)
	if not os.path.isdir(config):
		try:
			os.mkdir(config, 0700)
		except:
			parser.error("configuration directory '%s' does not exist, and could not be created" % options.config)
			sys.exit(1)

	gflags.FLAGS.auth_local_webserver = options.auth_local_webserver

	if not args:
		args = ["sendall"]

	credentials = None
	service = None

	i = 0
	while i < len(args):
		action = args[i]
		i += 1

		if action in ("auth", "authenticate", "authorize"):
			get_credentials(config, authorize = True)
		else:
			if credentials is None:
				credentials = get_credentials(config)
				service = get_service(credentials)

			if action == "list":
				for entry in get_all_calendars(service):
					print entry["summary"]
					print "  Id: " + entry["id"]
					print

			elif action == "subscribe":
				summary_or_id = args[i]
				i += 1

				subscriptions = get_subscriptions(config)
				for entry in get_all_calendars(service):
					if summary_or_id == entry["id"] or summary_or_id.lower() == entry["summary"].lower():
						subscriptions[entry["id"]] = { "summary": entry["summary"], "recipients": {} }
						set_subscriptions(config, subscriptions)
						break
				else:
					print >>sys.stderr, "Error: No calendar with Title or ID '%s' found" % summary_or_id

			elif action == "unsubscribe":
				summary_or_id = args[i]
				i += 1

				subscriptions = get_subscriptions(config)
				id, subscription = find_subscription(subscriptions, summary_or_id)
				if id:
					del subscriptions[id]
					set_subscriptions(config, subscriptions)

			elif action == "subscriptions":
				subscriptions = get_subscriptions(config)
				for id, subscription in subscriptions.iteritems():
					print subscription["summary"]
					print "  Id: " + id
					print "  Recipients: "
					for email, name in subscription["recipients"].iteritems():
						print "    %s <%s>" % (name, email)

					print "  Config:"
					for key, value in subscription.iteritems():
						if key in ("recipients", "summary"):
							continue
						print "    %s: %s" % (key, value)

					print

			elif action == "add":
				summary_or_id = args[i]
				i += 1
				email = args[i]
				i += 1
				name = args[i]
				i += 1

				subscriptions = get_subscriptions(config)
				id, subscription = find_subscription(subscriptions, summary_or_id)
				if id:
					subscription["recipients"][email] = name
					set_subscriptions(config, subscriptions)

			elif action in ("remove", "del"):
				summary_or_id = args[i]
				i += 1
				email_or_name = args[i]
				i += 1

				subscriptions = get_subscriptions(config)
				id, subscription = find_subscription(subscriptions, summary_or_id)
				if id:
					for email, name in subscription["recipients"].iteritems():
						if email == email_or_name or name == email_or_name:
							del subscription["recipients"][email]
							set_subscriptions(config, subscriptions)
					else:
						print >>sys.stderr, "Error: no recipient with email or name '%s' found" % email_or_name

			elif action == "set":
				summary_or_id = args[i]
				i += 1
				name = args[i]
				i += 1
				value = args[i]
				i += 1

				subscriptions = get_subscriptions(config)
				id, subscription = find_subscription(subscriptions, summary_or_id)
				if id:
					subscription[name] = value
					set_subscriptions(config, subscriptions)

			elif action == "sendall":
				subscriptions = get_subscriptions(config)
				for id, subscription in subscriptions.iteritems():
					events = get_events(service, id, subscription)
					send_email(config, subscription, events, options.dryrun)

			elif action == "send":
				summary_or_id = args[i]
				i += 1

				subscriptions = get_subscriptions(config)
				id, subscription = find_subscription(subscriptions, summary_or_id)
				if id:
					events = get_events(service, id, subscription)
					send_email(config, subscription, events, options.dryrun)

			else:
				parser.error("unknown action '%s'" % action)

	sys.exit(0)

if __name__ == "__main__":
	main()