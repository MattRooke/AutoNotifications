#!/usr/bin/python
from datetime import datetime as date
from O365 import *
import printer
import requests
import json
from openpyxl import load_workbook, cell

version = 0.6
authors = "Matthew Rooke, Ryan Clark"

email_repot_body = "<!DOCTYPE html><head><link rel=\"stylesheet\" type=\"text/css\"/><style type=\"text/css\">body{font-family: sans-serif;}.fail{background-color: red !important;}.success{background-color: greenyellow !important;}.alert{background-color: yellow !important;}<style/></head><body><div class=\"container-fluid\"><img class=\"img-fluid\" src=\"https://company.com.au/images/logo.jpg\" /><h1>AutoNotifications : Full Report</h1></div>"

printer.print_title(version, authors)
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

if not account.is_authenticated:
    print(
        "\n╔═════════════════════════╗\n"
        "║ Complete Authentication ║\n"
        "╚═════════════════════════╝"
    )
    account.authenticate()

mailbox = account.mailbox()
inbox = mailbox.inbox_folder()

# get current customer folders;
#print(
#    "╔══════════════════════════╗\n"
#    "║ Loading customer folders ║\n"
#    "╠══════════════════════════╝"
#)
#status_dictionary = {}
#customers_dictionary = {}
#customers = inbox.get_folder(folder_name="Customers")
#customer_folders = customers.get_folders()
#for folder in customer_folders:
#    print("║ " + folder.name)
#    customers_dictionary["{0}".format(folder.name)] = customers.get_folder(
#        folder_name="{0}".format(folder.name)
#    )
#    status_dictionary["{0}".format(folder.name)] = ""
#print("╠══════════════════════════\n" "║\n" "║")


def sort_for_inbox(customer_name, key_word, backup_type):
    # TODO use multiple keywords***
    global status_dictionary
    global email_repot_body
    query = inbox.new_query()
    query = query.on_attribute("subject").contains(key_word)
    query = query.chain("or").on_attribute("from").contains(key_word)
    filtered_messages = inbox.get_messages(query=query, limit=10000, batch=100)
    # TODO ignore empty queries
    count = 0
    status = ""
    print("╠════════\n" "║ Sorting - {} \n" "╠═".format(customer_name))
    email_repot_body += "<divclass=\"container-fluid\"><h3>Sorted Mail - {} </h3><table class=\"table\"><thead></thead><tbody>".format(customer_name)
    for email in filtered_messages:
        email.mark_as_read()
        email.move(folder=customers_dictionary[customer_name])
        print("║ {}".format(email))
        email_repot_body += "<tr><td scope=\"row\">{}</td></tr>".format(email)
        count += 1
        # Status codes: "NB": No Backup, "NS": NAS Success, "NF": NAS Failure,"AS" AFI Backup Success, "AF" AFI Backup Fail, "BS": Backup Success, "BF": Backup Failure.
        subject_line = "{}".format(email).lower()
        if "nas" in subject_line and backup_type != "none":
            if "success" in subject_line:
                status += "NS,"
            else:
                status += "NF,"
        elif (
            "finished" in subject_line
            or "success" in subject_line
            or "good" in subject_line
            or "completed" in subject_line
        ):
            status += "BS,"
        #AFI backup emails are only recived on a sucessful backup
        elif "afi" in subject_line:
            status += "AS,"
        elif "fail" in subject_line or "bad" in subject_line and backup_type != "NAS":
            status += "BF,"
    if backup_type == "none":
        status += "NB,"
    if count == 0 and backup_type == "afi":
        status += "AF"
    status_dictionary["{0}".format(customer_name)] = (
        status_dictionary["{0}".format(customer_name)] + status
    )
    email_repot_body += "</tbody></table><h4>Sorted {} emails for {}</h4></div>".format(
        count, customer_name)

    print(
        "╠═\n"
        "║ Sorted {} emails for {} \n"
        "╠════════\n"
        "║".format(count, customer_name)
    )


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
    print("╠═══════════════╗\n" "║ Status Report ║\n" "╠═══════════════╝\n" "║")
    email_repot_body += "<div class=\"container-fluid\"><h2>Quick Status Report</h2><table class=\"table\"><thead></thead><tbody>"
    report = ""
    for customer in customer_folders:

        status = status_dictionary["{}".format(customer.name)]

        if "NB" in status:
            print("╠════\n" "║ {} has no sceduled backup.".format(customer.name))
            report = report + ("{} has no sceduled backup.<br>".format(customer.name))
        if "AF" in status:
            print("╠════\n" "║ {} AFI backup : fail.".format(customer.name))
            email_repot_body += "<tr><td class=\"fail\" scope=\"row\">{} AFI backup : fail.</td></tr>".format(customer.name)
            report = report + ("{} AFI backup : fail <br>".format(customer.name))
        elif "AS" in status:
            print("╠════\n" "║ {} AFI backup : success.".format(customer.name))
            email_repot_body += "<tr><td class=\"success\" scope=\"row\">{} AFI backup : success.</td></tr>".format(customer.name)
            report = report + ("{} AFI backup : success <br>".format(customer.name))
        if "NF" in status:
            print("╠════\n" "║ {} NAS : fail".format(customer.name))
            email_repot_body += "<tr><td class=\"fail\" scope=\"row\">{} NAS : fail.</td></tr>".format(customer.name)
            report = report + ("{} NAS : fail <br>".format(customer.name))
        elif "NS" in status:
            print("╠════\n" "║ {} NAS : success".format(customer.name))
            email_repot_body += "<tr><td class=\"success\" scope=\"row\">{} NAS : success.</td></tr>".format(customer.name)
            report = report + ("{} NAS : success <br>".format(customer.name))
        if "BF" in status:
            print("╠════\n" "║ {} Backup : fail".format(customer.name))
            email_repot_body += "<tr><td class=\"fail\" scope=\"row\">{} Backup : fail.</td></tr>".format(customer.name)
            report = report + ("{} Backup : fail <br>".format(customer.name))
        elif "BS" in status:
            print("╠════\n" "║ {} Backup : success".format(customer.name))
            email_repot_body += "<tr><td class=\"success\" scope=\"row\">{} Backup : success.</td></tr>".format(customer.name)
            report = report + \
                ("{} Backup : success <br>".format(customer.name))
        if status == '':
            print(
                "╠════\n" "║ {} status : no status emails received".format(
                    customer.name)
            )
            email_repot_body += "<tr><td class=\"alert\" scope=\"row\">{} : No status emails received.</td></tr>".format(customer.name)
            report = report + (
                "{} status : no status email received <br>".format(
                    customer.name)
            )
    email_repot_body += "</tbody></table></div>"
    email_repot_body += "<p>Debug - Status:{}</p>".format(status_dictionary)
    return report


def email_report():
    global email_repot_body
    msg = account.new_message()
    msg.to.add("matt@company.com.au")
    msg.subject = "Notification Report " + date.today().strftime("%d/%m/%Y")
    msg.body = email_repot_body
    msg.send()
    print("╠════\n" "║ Email report sent successfully")


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


def update_spreadsheet():
    FILE_PATH = 'T:\True North Technology Team Site - Documents\Backup Check - Copy.xlsx'
    workbook = load_workbook(FILE_PATH)
    ws = workbook.active
    current = ""
    for i in ws.iter_rows(max_row=100, max_col=1):
        for cell in i:
            if cell.value == "Customer Name/Computer Name":  
                date_row = cell.row  #Assigns the value of the row which contains the dates 
            if cell.value == "Restore Test":
                number_of_rows = cell.row - date_row
                break
    for cell in ws[date_row]:  #Finds the empty cell which will use today's date
        if cell.value == None:
            ws['C1'] = "Test"
            todays_column = cell.column
            ws.cell(row=cell.row, column=cell.column).value = 1
            print(cell.value)
            print(ws.cell.value)
            cell.value = date.today().strftime("%d/%m/%Y")
            current = cell.column
            break
    workbook.save

                    #if col.value == None:
                    #    col.value = date.today().strftime("%d/%m/%Y")
                    #    current = col.column
                    #    break

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
        filename=FILE_PATH
    )


update_spreadsheet()

print("╠═══════════════╗\n" "║ Reading Inbox ║\n" "╠═══════════════╝\n" "║")

# TODO Read and react to alerts!

# Robot list:
email_repot_body += "<h2>Robot emails:</h2>"
#sort_for_inbox("Atera Alerts", "atera", "none")
sort_for_inbox("Uptime Robot", "uptimerobot", "none")

# Customer list:
# Backup types: none, afi = afi.ai cloud, SPX = Shadow Protect, Veeam, NAS, , C2 = Code2, Other
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
email_repot_body += "<h2>company interanl emails:</h2>"
sort_for_inbox("company", "Microsoft", "other")
sort_for_inbox("company", "All Customer", "other")
sort_for_inbox("company", "company", "Veeam")
sort_for_inbox("company", "usecure", "other")
sort_for_inbox("company", "dmarc", "other")
sort_for_inbox("company", "IT Automation Task Feedback", "none")
sort_for_inbox("company", "Alert; Problem", "none")
sort_for_inbox("company", "BitDefender", "none")

query_notification_status()

#get_atera_alerts()

# TODO Sort all other emails | Other folder?

# Debug:
print("Debug:")
# print("Nothing to see here...")
print(status_dictionary)

# email
email_repot_body += "</body>"


print("║\n" "╠══════════╗\n" "║ Complete ║\n" "╚══════════╝")

update_spreadsheet()
email_report()
