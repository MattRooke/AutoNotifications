#!/usr/bin/python
from datetime import datetime as date
from O365 import *
import requests
import json
#from openpyxl import load_workbook, cell

#Debug``
input("Press any Key to start...")
print('Working...')

version = 0.6
authors = "Matthew Rooke, Ryan Clark"

email_repot_body = "<!DOCTYPE html><head><link rel=\"stylesheet\" type=\"text/css\"/><style type=\"text/css\">body{font-family: sans-serif;}.fail{background-color: red !important;}.success{background-color: greenyellow !important;}.alert{background-color: yellow !important;}<style/></head><body><div class=\"container-fluid\"><img class=\"img-fluid\" src=\"https://company.com.au/images/logo.jpg\" /><h1>AutoNotifications : Full Report</h1></div>"

credentials = (
    "...",
)

protocols = MSGraphProtocol(
    default_resource="notifications@company.com.au")
scopes = protocols.get_scopes_for(["message_all_shared", "offline_access"])
token_backend = FileSystemTokenBackend(
    token_path="my_folder", token_filename="my_token.txt"
)
account = Account(
    credentials, scopes=scopes, protocol=protocols, token_backend=token_backend
)
con = Connection(credentials, scopes=scopes)

if not account.is_authenticated:
    account.authenticate()

mailbox = account.mailbox()
inbox = mailbox.inbox_folder()

# get current customer folders;
status_dictionary = {}
customers_dictionary = {}
customers = inbox.get_folder(folder_name="Customers")
customer_folders = customers.get_folders()
for folder in customer_folders:
    customers_dictionary["{0}".format(folder.name)] = customers.get_folder(
        folder_name="{0}".format(folder.name)
    )
    status_dictionary["{0}".format(folder.name)] = ""


def sort_for_inbox(customer_name, key_word):
    # TODO use multiple keywords***
    global status_dictionary
    global email_repot_body
    query = inbox.new_query()
    query = query.on_attribute("subject").contains(key_word)
    query = query.chain("or").on_attribute("from").contains(key_word)
    filtered_messages = inbox.get_messages(query=query, limit=10000, batch=100)
    # TODO ignore empty queries > TEST
    count = 0
    status = ""
    length = 0
    # TODO print to .txt file and email
    email_repot_body += "<divclass=\"container-fluid\"><h2>Sorted Mail - {} </h2><table class=\"table\"><thead></thead><tbody>".format(customer_name)
    for email in filtered_messages:
        email.mark_as_read()
        email.move(folder=customers_dictionary[customer_name])
        email_repot_body += "<tr><td scope=\"row\">{}</td></tr>".format(email)
        count += 1
        # Status codes: "NS": NAS Success, "NF": NAS Failure, "BS": Backup Success, "BF": Backup Failure.
        subject_line = "{}".format(email).lower()
        if "nas" in subject_line:
            if "success" in subject_line:
                status += "NS,"
            elif "fail" in subject_line:
                status += "NF,"
        elif (
            "finished" in subject_line
            or "success" in subject_line
            or "good" in subject_line
        ):
            status += "BS,"
        elif "fail" in subject_line or "bad" in subject_line:
            status += "BF,"
    status_dictionary["{0}".format(customer_name)] = (
        status_dictionary["{0}".format(customer_name)] + status
    )
    email_repot_body += "</tbody></table><h3>Sorted {} emails for {}</h3></div>".format(
        count, customer_name)


def query_sub_folders_for_customers(customer_name):
    sub_folder_dictionary = {}
    sub_folders = customers_dictionary[customer_name].get_folders()
    for sub_folder in sub_folders:
        sub_folder_dictionary["{0}".format(sub_folder.name)] = customers.get_folder(
            folder_name="{0}".format(sub_folder.name)
        )
    return sub_folder_dictionary


def query_notification_status():
    # TODO show report with email list
    global email_repot_body
    email_repot_body += "<div class=\"container-fluid\"><h2>Quick Status Report</h2><table class=\"table\"><thead></thead><tbody>"
    report = ""
    for customer in customer_folders:

        status = status_dictionary["{}".format(customer.name)]

        if "NF" in status:
            email_repot_body += "<tr><td class=\"fail\" scope=\"row\">{} NAS : fail.</td></tr>".format(customer.name)
            report = report + ("{} NAS : fail <br>".format(customer.name))
        elif "NS" in status:
            email_repot_body += "<tr><td class=\"success\" scope=\"row\">{} NAS : success.</td></tr>".format(customer.name)
            report = report + ("{} NAS : success <br>".format(customer.name))
        if "BF" in status:
            email_repot_body += "<tr><td class=\"fail\" scope=\"row\">{} Backup : fail.</td></tr>".format(customer.name)
            report = report + ("{} Backup : fail <br>".format(customer.name))
        elif "BS" in status:
            email_repot_body += "<tr><td class=\"success\" scope=\"row\">{} Backup : success.</td></tr>".format(customer.name)
            report = report + \
                ("{} Backup : success <br>".format(customer.name))
        else:
            email_repot_body += "<tr><td class=\"alert\" scope=\"row\">{} : No status emails received.</td></tr>".format(customer.name)
            report = report + (
                "{} status : no status email received <br>".format(
                    customer.name)
            )
    email_repot_body += "</tbody></table></div>"
    return report


def email_report():
    global email_repot_body
    msg = account.new_message()
    msg.to.add("ryan@company.com.au")
    msg.to.add("matt@company.com.au")
    msg.subject = "TEST Notification Report " + date.today().strftime("%d/%m/%Y")
    msg.body = email_repot_body
    msg.send()
    print("Email report sent successfully")


def get_atera_alerts():
    global email_repot_body
    email_repot_body += "<div class=\"container-fluid\"><h2>Atera Alerts</h2><table class=\"table\"><thead></thead><tbody>"
    headers = {
        "Accept": "application/json",
        "X-API-KEY": "...",
    }
    response = requests.get(
        "https://app.atera.com/api/v3/alerts?itemsInPage=1000", headers=headers
    )
    jsondata = json.loads(response.text)
    criticaldata = {}
    for alert in jsondata["items"]:
        if alert["Severity"] == "Critical":
            criticaldata[alert["DeviceName"]] = alert["AlertMessage"]
    num = 1
    for entry in criticaldata:
        email_repot_body += "<tr><td scope=\"row\">{}. {} <br>{}<br></td></tr>".format(num, entry, criticaldata.get(entry))
        num += 1
    email_repot_body += "</tbody></table></div>"


def update_spreadsheet():
    workbook = load_workbook(
        filename=R"C:\Users\technician1\company\company Team Site - Documents\test.xlsx"
    )
    ws = workbook.active
    current = ""
    for i in ws.iter_rows(max_row=1, max_col=1000):
        for col in i:
            if col.value == None:
                col.value = date.today().strftime("%d/%m/%Y")
                current = col.column
                break

    print(current)
    notification_status = query_notification_status().split("<br>")
    notification_dict = {}
    """
    for x in notification_status:
        temp = x.split(" : ")
        print(temp)
        notification_dict[temp[0]] = temp[1]
    
    """
    for x in ws.iter_cols(
        min_col=int(current), max_col=int(current), max_row=31, min_row=2
    ):
        for row in x:
            row.style

    workbook.save(
        filename=R"C:\Users\technician1\company\company Team Site - Documents\test.xlsx"
    )


print('Sorting...')

# TODO Read and react to alerts!

# Robot list:
sort_for_inbox("Atera Alerts", "atera")
sort_for_inbox("Uptime Robot", "uptimerobot")

# Customer list:
# sort_for_inbox("...", "...")
# sort_for_inbox("...", "...")
# sort_for_inbox("...", "...")
# sort_for_inbox("...", "...")
# sort_for_inbox("...", "...")
# sort_for_inbox("...", "...")
# sort_for_inbox("...", "...")
# sort_for_inbox("...", "...")
# sort_for_inbox("...", "...")
# sort_for_inbox("...", "...")
# sort_for_inbox("...", "...")
# sort_for_inbox("...", "...")
# sort_for_inbox("...", "...")
# sort_for_inbox("...", "...")
# sort_for_inbox("...", "...")
# sort_for_inbox("...", "...")
# sort_for_inbox("...", "...")
# sort_for_inbox("...", "...")
# sort_for_inbox("...", "...")
# sort_for_inbox("...", "...")
# sort_for_inbox("...", "...")
# sort_for_inbox("...", "...")
... 
# sort_for_inbox("...", "...")

query_notification_status()

get_atera_alerts()

# TODO Sort all other emails | Other folder?

# Debug:
print("Debug:")
input('Press any key to end...')
#print("Nothing to see here...")
# print(status_dictionary)

# email
email_repot_body += "</body>"
email_report()

print('Done!')
# update_spreadsheet()
