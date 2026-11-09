#!/usr/bin/python
from datetime import datetime as date
from O365 import *
import requests
import json
#from openpyxl import load_workbook, cell


email_repot_body = "<!DOCTYPE html><head><link rel=\"stylesheet\" type=\"text/css\"/><style type=\"text/css\">body{font-family: sans-serif;}.fail{background-color: red !important;}.success{background-color: greenyellow !important;}.alert{background-color: yellow !important;}<style/></head><body><div class=\"container-fluid\"><img class=\"img-fluid\" src=\"https://company.com.au/wp-content/uploads/2022/07/Logo-White-No-Tagline-1024x285.png\" /><h1>AutoNotifications : Full Report</h1></div>"

credentials = (
    "...",
)

protocols = MSGraphProtocol(
    default_resource="notifications@company.com.au")
scopes = protocols.get_scopes_for(["message_all_shared", "offline_access"])
token_backend = FileSystemTokenBackend(
    token_path="tokens", token_filename="main_token.txt"
)
account = Account(
    credentials, scopes=scopes, protocol=protocols, token_backend=token_backend
)
con = Connection(credentials, scopes=scopes)


mailbox = account.mailbox()
inbox = mailbox.inbox_folder()

# load current customer folders;
status_dictionary = {}
customers_dictionary = {}
customers = inbox.get_folder(folder_name="Customers")
customer_folders = customers.get_folders()
for folder in customer_folders:
    customers_dictionary["{0}".format(folder.name)] = customers.get_folder(
        folder_name="{0}".format(folder.name)
    )
    status_dictionary["{0}".format(folder.name)] = ""


def sort_for_inbox(customer_name, key_word, backup_type):
    # TODO use multiple keywords***
    global status_dictionary
    global email_repot_body
    query = inbox.new_query()
    query = query.on_attribute("subject").contains(key_word)
    query = query.chain("or").on_attribute("from").contains(key_word)
    filtered_messages = inbox.get_messages(query=query, limit=10000, batch=1000)
    # TODO ignore empty queries
    count = 0
    status = ""

    email_repot_body += "<divclass=\"container-fluid\"><h3>Sorted Mail - {} </h3><table class=\"table\"><thead></thead><tbody>".format(customer_name)
    for email in filtered_messages:
        email.mark_as_read()
        email.move(folder=customers_dictionary[customer_name])
        email_repot_body += "<tr><td scope=\"row\">{}</td></tr>".format(email)
        count += 1

        # Status codes: "NB": No Backup, "NS": NAS Success, "NF": NAS Failure,"AS" AFI Backup Success, "AF" AFI Backup Fail, "BS": Backup Success, "BF": Backup Failure.
        subject_line = "{}".format(email).lower()
        if "nas" in subject_line and backup_type != "none":
            if (
                "success" in subject_line 
                or "completed" in subject_line
                ):
                status += "NS,"
            else:
                status += "NF,"
        # Success keywords in subject line
        elif (
            "finished" in subject_line
            or "success" in subject_line
            or "good" in subject_line
            or "completed" in subject_line
        ):
            status += "BS,"
        
        #AFI backup emails are only recived on a sucessful backup
        elif "afi" in subject_line:
            status += "AFS,"
        elif "fail" in subject_line or "bad" in subject_line and backup_type != "NAS":
            status += "BF," 
    if backup_type == "none":
        status += "NB,"
    if count == 0 and backup_type == "afi":
        status += "AFF,"
    if "NAS" in backup_type:
        status += "NAS,"
    if "BACKUP" in backup_type or "afi" in backup_type:
        status += "BUP,"

    status_dictionary["{0}".format(customer_name)] = (
        status_dictionary["{0}".format(customer_name)] + status
    )
    email_repot_body += "</tbody></table><h4>Sorted {} emails for {}</h4></div>".format(
        count, customer_name)



def query_sub_folders_for_customers(customer_name):
    sub_folder_dictionary = {}
    sub_folders = customers_dictionary[customer_name].get_folders()
    for sub_folder in sub_folders:
        print(sub_folder.name)
        sub_folder_dictionary["{0}".format(sub_folder.name)] = customers.get_folder(
            folder_name="{0}".format(sub_folder.name)
        )
    return sub_folder_dictionary

def query_notification_status():
    global email_repot_body
    email_repot_body += "<div class=\"container-fluid\"><h2>Quick Status Report</h2><h3>Backup:</h3><table class=\"table\"><thead></thead><tbody>"
    report = ""
    for customer in customer_folders:

        status = status_dictionary["{}".format(customer.name)]

        if "BUP" in status:
            if "NB" in status:
                print("╠════\n" "║ {} has no monitored backup.".format(customer.name))
            if "AFF" in status:
                print("╠════\n" "║ {} AFI backup : fail.".format(customer.name))
                email_repot_body += "<tr><td class=\"fail\" scope=\"row\">{} AFI : fail.</td></tr>".format(customer.name)
            elif "AFS" in status:
                print("╠════\n" "║ {} AFI backup : success.".format(customer.name))
                email_repot_body += "<tr><td class=\"success\" scope=\"row\">{} AFI : success.</td></tr>".format(customer.name)
            elif "BF" in status:
                print("╠════\n" "║ {} : fail".format(customer.name))
                email_repot_body += "<tr><td class=\"fail\" scope=\"row\">{} : fail.</td></tr>".format(customer.name)
            elif "BS" in status:
                print("╠════\n" "║ {} : success".format(customer.name))
                email_repot_body += "<tr><td class=\"success\" scope=\"row\">{} : success.</td></tr>".format(customer.name)
            else:
                email_repot_body += "<tr><td class=\"alert\" scope=\"row\">{} : Warnings or no emails received.</td></tr>".format(customer.name)

    email_repot_body += "</tbody></table></div>"

    email_repot_body += "<div class=\"container-fluid\"><h3>NAS:</h3><table class=\"table\"><thead></thead><tbody>"
    report = ""
    for customer in customer_folders:

        status = status_dictionary["{}".format(customer.name)]

        if "NAS" in status:
            if "NF" in status:
                print("╠════\n" "║ {} NAS : fail".format(customer.name))
                email_repot_body += "<tr><td class=\"fail\" scope=\"row\">{} : fail.</td></tr>".format(customer.name)
            elif "NS" in status:
                print("╠════\n" "║ {} NAS : success".format(customer.name))
                email_repot_body += "<tr><td class=\"success\" scope=\"row\">{} : success.</td></tr>".format(customer.name)
            else:
                email_repot_body += "<tr><td class=\"alert\" scope=\"row\">{} : Warnings or no emails received.</td></tr>".format(customer.name)


    email_repot_body += "</tbody></table></div>"

    # email_repot_body += "<p>Debug - Status:{}</p>".format(status_dictionary)
    return report

# Email Report sending add poeple to send to ect
def email_report():
    global email_repot_body
    msg = account.new_message()
    msg.to.add("adam@company.com.au")
    msg.to.add("sheena@company.com.au")
    msg.to.add("jay@company.com.au")
    msg.to.add("oliver@company.com.au")
    msg.to.add("steve@company.com.au")
    msg.to.add("12345678.company.onmicrosoft.com@apac.teams.ms")
    msg.subject = "Notification Report " + date.today().strftime("%d/%m/%Y")
    msg.body = email_repot_body
    msg.send()


# Dont think this functions
def get_atera_alerts():
    global email_repot_body
    print("╠══════════════╗\n║ Atera Alerts ║\n╠══════════════╝\n║")
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
        print("╠════\n" "║{}. {} \n" "║{}".format(
            num, entry, criticaldata.get(entry)))
        email_repot_body += "<tr><td scope=\"row\">{}. {} <br>{}<br></td></tr>".format(num, entry, criticaldata.get(entry))
        num += 1
    email_repot_body += "</tbody></table></div>"

#TODO Update Spreadsheet automatically
# def update_spreadsheet():
#     workbook = load_workbook(
#         filename=R"C:\Users\technician1\company\company Team Site - Documents\test.xlsx"
#     )
#     ws = workbook.active
#     current = ""
#     for i in ws.iter_rows(max_row=1, max_col=1000):
#         for col in i:
#             if col.value == None:
#                 col.value = date.today().strftime("%d/%m/%Y")
#                 current = col.column
#                 break

#     print(current)
#     notification_status = query_notification_status().split("<br>")
#     notification_dict = {}
#     """
#     for x in notification_status:
#         temp = x.split(" : ")
#         print(temp)
#         notification_dict[temp[0]] = temp[1]
    
#     """
#     for x in ws.iter_cols(
#         min_col=int(current), max_col=int(current), max_row=31, min_row=2
#     ):
#         for row in x:
#             row.style

#     workbook.save(
#         filename=R"C:\Users\technician1\company\company Team Site - Documents\test.xlsx"
#     )


# Robot list:
email_repot_body += "<h2>Robot emails:</h2>"
#sort_for_inbox("Atera Alerts", "atera", "none")
sort_for_inbox("Uptime Robot", "uptimerobot", "none")

# Customer list:
# Backup types: none, afi = afi.ai cloud, BACKUP = Shadow Protect & Veeam & Code2, NAS, Other
# Passes keywords and customer name through the sort_for_inbox
email_repot_body += "<h2>Customer emails:</h2>"
sort_for_inbox("...", "...")
sort_for_inbox("...", "...")
sort_for_inbox("...", "...")
sort_for_inbox("...", "...")
sort_for_inbox("...", "...")
sort_for_inbox("...", "...")
sort_for_inbox("...", "...")
sort_for_inbox("...", "...")
sort_for_inbox("...", "...")
sort_for_inbox("...", "...")
sort_for_inbox("...", "...")
sort_for_inbox("...", "...")
sort_for_inbox("...", "...")
sort_for_inbox("...", "...")
sort_for_inbox("...", "...")
sort_for_inbox("...", "...")
sort_for_inbox("...", "...")
sort_for_inbox("...", "...")
sort_for_inbox("...", "...")
sort_for_inbox("...", "...")
sort_for_inbox("...", "...")
sort_for_inbox("...", "...")
... 
sort_for_inbox("...", "...")

email_repot_body += "<h2>company internal emails:</h2>"
sort_for_inbox("company", "Atera", "other")
sort_for_inbox("company", "Microsoft", "other")
sort_for_inbox("company", "All Customer", "other")
sort_for_inbox("company", "company", "BACKUP")
sort_for_inbox("company", "usecure", "other")
sort_for_inbox("company", "dmarc", "other")
sort_for_inbox("company", "IT Automation Task Feedback", "none")
sort_for_inbox("company", "Alert;", "none")
sort_for_inbox("company", "BitDefender", "none")
sort_for_inbox("company", "trendmicro", "none")
sort_for_inbox("company", "letsencrypt", "none")
sort_for_inbox("company", "Support Request Received", "none")

query_notification_status()

#get_atera_alerts()

# TODO Sort all other emails | Other folder?

# Debug:
print("Debug:")
print("Nothing to see here...")
# print(status_dictionary)

# email
email_repot_body += "</body>"
email_report()


print("║\n" "╠══════════╗\n" "║ Complete ║\n" "╚══════════╝")
# update_spreadsheet()
