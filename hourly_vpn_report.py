# Adding Comment to test Git for Version Controlling
import datetime
import pandas as pd
import xlwings as wings
import os
import os.path
import email
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import glob
import time
from pytz import timezone
import gc
import keyring
from O365 import Account
day_name = datetime.datetime.now().strftime("%A")
if day_name in ['Saturday', 'Sunday']:
    exit(2001)


"""
Initialization of Parameters i.e. define path of Source & Destination
"""
bcp_output_path = "D:\\Python\\ICT_VPN_Report\\BCP_REPORT\\Report_To_Send\\"
bcp_report_name = "BCP_Internet_VPN_Report-"
curr_dated_file = datetime.datetime.now().strftime("%d_%B_%Y") + ".xlsx"
bcp_report_name = bcp_report_name + curr_dated_file
bcp_report_filepath = bcp_output_path + bcp_report_name
solar_winds_path = "D:\\Python\\ICT_VPN_Report\\BCP_REPORT\\Solar_Winds\\"
calculation_filepath = str(bcp_output_path) + "DnD_Calculation\\" + "Calculation_Method.xlsx"
curr_hour = datetime.datetime.now().strftime("%H")
excel_max_row = 1048576
app = wings.App(visible=False)
solar_winds_subject = "FW: Internet/Wan Utilization - Internal IT Networks - Peak BPS Report"
vpn_subject = "FW: VPN Hourly report"


def download_bcp_report():
    pws_graph = keyring.get_password('Graph_api', 'rpa_user04')
    user, pass1, tid = pws_graph.split('|')[0], pws_graph.split('|')[1], pws_graph.split('|')[2]
    credentials = (user, pass1)
    account = Account(credentials, auth_flow_type='credentials', tenant_id=tid)
    if account.authenticate():
        mailbox = account.mailbox('RPA_USER04@colt.net')
        for message in mailbox.get_messages(download_attachments=True):
            time.sleep(.5)
            if message.subject == solar_winds_subject:
                for attachment in message.attachments:
                    attachment_name = attachment.name
                    if attachment.name.split(".")[-1] == "csv":
                        attachment.save(solar_winds_path)
                        break
        df = pd.read_csv(solar_winds_path + attachment_name, encoding='utf-8')
        mailbox = account.mailbox('RPA_USER04@colt.net')
        for message in mailbox.get_messages(download_attachments=True):
            time.sleep(.5)
            if message.subject == vpn_subject:
                raw_data = pd.read_html(message.body)
                cyber_vpn_report = raw_data[0].drop(columns=["Unnamed: 0"])
                wb = wings.Book(calculation_filepath)
                ws = wb.sheets[0]
                ws.autofit()
                ws.range('B3').value = cyber_vpn_report
                wb.save(calculation_filepath)
                wb.close()
        return df


if __name__ == "__main__":
    solar_winds = download_bcp_report()
    new_sw_data = solar_winds[['Interface Alias', 'Peak Receive bps', 'Peak Transmit bps', 'Full Name']]
    new_sw_data.apply(lambda x: x.astype(str).str.upper())
    internet = new_sw_data[new_sw_data['Full Name'].str.contains("-INT")]
    ggn_internet = internet[internet['Full Name'].str.contains("GGN")]
    ggn_internet = ggn_internet.reset_index(drop=True)
    blr_internet = internet[internet['Full Name'].str.contains("BLR")]
    blr_internet = blr_internet.reset_index(drop=True)
    mpls = new_sw_data[new_sw_data['Full Name'].str.contains("-W")]
    ggn_mpls = mpls[mpls['Full Name'].str.contains("GGN")]
    ggn_mpls = ggn_mpls.reset_index(drop=True)
    blr_mpls = mpls[mpls['Full Name'].str.contains("BLR")]
    blr_mpls = blr_mpls.reset_index(drop=True)
    eu_int = new_sw_data[new_sw_data['Full Name'].str.contains("LON/")]
    eu_int = eu_int.reset_index(drop=True)
    ggn_int_list = []
    for i in range(ggn_internet.shape[0]):
        peak_rec = round(float(ggn_internet["Peak Receive bps"][i].split()[0]))
        if ggn_internet["Peak Receive bps"][i].split()[1].upper() == "BPS":
            peak_rec = round(peak_rec / 1048576)
        elif ggn_internet["Peak Receive bps"][i].split()[1].upper() == "KBPS":
            peak_rec = round(peak_rec / 1024)
        ggn_int_list.append(peak_rec)
        peak_transit = round(float(ggn_internet["Peak Transmit bps"][i].split()[0]))
        if ggn_internet["Peak Transmit bps"][i].split()[1].upper() == "BPS":
            peak_transit = round(peak_transit / 1048576)
        elif ggn_internet["Peak Transmit bps"][i].split()[1].upper() == "KBPS":
            peak_transit = round(peak_transit / 1024)
        ggn_int_list.append(peak_transit)
    max_ggn_int = max([int(item) for item in ggn_int_list])
    blr_int_list = []
    for i in range(blr_internet.shape[0]):
        peak_rec = round(float(blr_internet["Peak Receive bps"][i].split()[0]))
        if blr_internet["Peak Receive bps"][i].split()[1].upper() == "BPS":
            peak_rec = round(peak_rec / 1048576)
        elif blr_internet["Peak Receive bps"][i].split()[1].upper() == "KBPS":
            peak_rec = round(peak_rec / 1024)
        blr_int_list.append(peak_rec)
        peak_transit = round(float(blr_internet["Peak Transmit bps"][i].split()[0]))
        if blr_internet["Peak Transmit bps"][i].split()[1].upper() == "BPS":
            peak_transit = round(peak_transit / 1048576)
        elif blr_internet["Peak Transmit bps"][i].split()[1].upper() == "KBPS":
            peak_transit = round(peak_transit / 1024)
        blr_int_list.append(peak_transit)
    max_blr_int = max([int(item) for item in blr_int_list])
    ggn_mpls_list = []
    for i in range(ggn_mpls.shape[0]):
        peak_rec = round(float(ggn_mpls["Peak Receive bps"][i].split()[0]))
        if ggn_mpls["Peak Receive bps"][i].split()[1].upper() == "BPS":
            peak_rec = round(peak_rec / 1048576)
        elif ggn_mpls["Peak Receive bps"][i].split()[1].upper() == "KBPS":
            peak_rec = round(peak_rec / 1024)
        ggn_mpls_list.append(peak_rec)
        peak_transit = round(float(ggn_mpls["Peak Transmit bps"][i].split()[0]))
        if ggn_mpls["Peak Transmit bps"][i].split()[1].upper() == "BPS":
            peak_transit = round(peak_transit / 1048576)
        elif ggn_mpls["Peak Transmit bps"][i].split()[1].upper() == "KBPS":
            peak_transit = round(peak_transit / 1024)
        ggn_mpls_list.append(peak_transit)
    max_ggn_mpls = max([int(item) for item in ggn_mpls_list])
    blr_mpls_list = []
    for i in range(blr_mpls.shape[0]):
        peak_rec = round(float(blr_mpls["Peak Receive bps"][i].split()[0]))
        if blr_mpls["Peak Receive bps"][i].split()[1].upper() == "BPS":
            peak_rec = round(peak_rec / 1048576)
        elif blr_mpls["Peak Receive bps"][i].split()[1].upper() == "KBPS":
            peak_rec = round(peak_rec / 1024)
        blr_mpls_list.append(peak_rec)
        peak_transit = round(float(blr_mpls["Peak Transmit bps"][i].split()[0]))
        if blr_mpls["Peak Transmit bps"][i].split()[1].upper() == "BPS":
            peak_transit = round(peak_transit / 1048576)
        elif blr_mpls["Peak Transmit bps"][i].split()[1].upper() == "KBPS":
            peak_transit = round(peak_transit / 1024)
        blr_mpls_list.append(peak_transit)
    max_blr_mpls = max([int(item) for item in blr_mpls_list])
    eu_int_list = []
    for i in range(eu_int.shape[0]):
        peak_rec = round(float(eu_int["Peak Receive bps"][i].split()[0]))
        if eu_int["Peak Receive bps"][i].split()[1].upper() == "BPS":
            peak_rec = round(peak_rec / 1048576)
        elif eu_int["Peak Receive bps"][i].split()[1].upper() == "KBPS":
            peak_rec = round(peak_rec / 1024)
        eu_int_list.append(peak_rec)
        peak_transit = round(float(eu_int["Peak Transmit bps"][i].split()[0]))
        if eu_int["Peak Transmit bps"][i].split()[1].upper() == "BPS":
            peak_transit = round(peak_transit / 1048576)
        elif eu_int["Peak Transmit bps"][i].split()[1].upper() == "KBPS":
            peak_transit = round(peak_transit / 1024)
        eu_int_list.append(peak_transit)
    max_eu_int = max([int(item) for item in eu_int_list])
    wb1 = wings.Book(calculation_filepath)
    ws1 = wb1.sheets[0]
    ws1.autofit()

    wb2 = wings.Book(bcp_report_filepath)
    ws2 = wb2.sheets[0]
    ws2.autofit()

    india_last_row = wb2.sheets[0].range('A' + str(wb2.sheets[0].cells.last_cell.row)).end('up').row
    if india_last_row == 1:
        india_last_row = 4
    else:
        india_last_row = india_last_row + 1
    dst_cell_range1 = 'B' + str(india_last_row)
    dst_cell_range2 = 'C' + str(india_last_row)
    dst_cell_range3 = 'D' + str(india_last_row)
    dst_cell_range4 = 'E' + str(india_last_row)
    dst_cell_range5 = 'F' + str(india_last_row) + ':H' + str(india_last_row)
    dst_cell_range6 = 'I' + str(india_last_row)
    dst_cell_range7 = 'J' + str(india_last_row)
    dst_cell_range8 = 'K' + str(india_last_row)
    dst_cell_range9 = 'L' + str(india_last_row) + ':Q' + str(india_last_row)
    time_cell_range = 'A' + str(india_last_row)
    wb2.sheets[0].range(time_cell_range).value = str(curr_hour) + ":00"
    wb2.sheets[0].range(dst_cell_range1).value = str(max_ggn_int)
    wb2.sheets[0].range(dst_cell_range2).value = str(max_blr_int)
    wb2.sheets[0].range(dst_cell_range3).value = str(max_ggn_mpls)
    wb2.sheets[0].range(dst_cell_range4).value = str(max_blr_mpls)
    wb2.sheets[0].range(dst_cell_range5).value = wb1.sheets[0].range('C20:E20').options(numbers=int).value
    wb2.sheets[0].range(dst_cell_range6).value = str(int(wb1.sheets[0].range('F20').value * 100)) + "%"
    wb2.sheets[0].range(dst_cell_range7).value = str(int(wb1.sheets[0].range('F20').value * 100)) + "%"
    wb2.sheets[0].range(dst_cell_range8).value = str(int(wb1.sheets[0].range('F20').value * 100)) + "%"
    wb2.sheets[0].range(dst_cell_range9).value = wb1.sheets[0].range('I20:N20').options(numbers=int).value

    eu_last_row = wb2.sheets[1].range('A' + str(wb2.sheets[1].cells.last_cell.row)).end('up').row
    if eu_last_row == 1:
        eu_last_row = 4
    else:
        eu_last_row = eu_last_row + 1
    dst_cell_range0 = 'B' + str(eu_last_row)
    dst_cell_range1 = 'C' + str(eu_last_row) + ':F' + str(eu_last_row)
    dst_cell_range2 = 'G' + str(eu_last_row)
    dst_cell_range3 = 'H' + str(eu_last_row)
    dst_cell_range4 = 'I' + str(eu_last_row)
    dst_cell_range5 = 'J' + str(eu_last_row)
    dst_cell_range6 = 'K' + str(eu_last_row) + ':R' + str(eu_last_row)
    time_cell_range = 'A' + str(eu_last_row)
    wb2.sheets[1].range(time_cell_range).value = str(curr_hour) + ":00"
    wb2.sheets[1].range(dst_cell_range0).value = str(max_eu_int)
    wb2.sheets[1].range(dst_cell_range1).value = wb1.sheets[0].range('C28:F28').options(numbers=int).value
    wb2.sheets[1].range(dst_cell_range2).value = str(int(wb1.sheets[0].range('G28').value * 100)) + "%"
    wb2.sheets[1].range(dst_cell_range3).value = str(int(wb1.sheets[0].range('H28').value * 100)) + "%"
    wb2.sheets[1].range(dst_cell_range4).value = str(int(wb1.sheets[0].range('I28').value * 100)) + "%"
    wb2.sheets[1].range(dst_cell_range5).value = str(int(wb1.sheets[0].range('J28').value * 100)) + "%"
    wb2.sheets[1].range(dst_cell_range6).value = wb1.sheets[0].range('K28:R28').options(numbers=int).value

    asia_last_row = wb2.sheets[2].range('A' + str(wb2.sheets[2].cells.last_cell.row)).end('up').row
    if asia_last_row == 1:
        asia_last_row = 4
    else:
        asia_last_row = asia_last_row + 1
    dst_cell_range1 = 'B' + str(asia_last_row) + ':C' + str(asia_last_row)
    dst_cell_range2 = 'D' + str(asia_last_row)
    dst_cell_range3 = 'E' + str(asia_last_row)
    dst_cell_range4 = 'F' + str(asia_last_row) + ':I' + str(asia_last_row)
    time_cell_range = 'A' + str(asia_last_row)

    now_utc = datetime.datetime.now(timezone('UTC'))
    now_asia = now_utc.astimezone(timezone('Asia/Tokyo'))
    curr_asia_hour = now_asia.strftime("%H")

    wb2.sheets[2].range(time_cell_range).value = str(curr_asia_hour) + ":30"
    wb2.sheets[2].range(dst_cell_range1).value = wb1.sheets[0].range('C34:D34').options(numbers=int).value
    wb2.sheets[2].range(dst_cell_range2).value = str(int(wb1.sheets[0].range('E34').value * 100)) + "%"
    wb2.sheets[2].range(dst_cell_range3).value = str(int(wb1.sheets[0].range('F34').value * 100)) + "%"
    wb2.sheets[2].range(dst_cell_range4).value = wb1.sheets[0].range('G34:J34').options(numbers=int).value

    wb2.save(bcp_report_filepath)
    wb2.close()
    app.quit()

    """
    Sending Report as in email attachment to stakeholder  
    """
    time.sleep(5)
    curr_time = datetime.datetime.now().strftime("%H:00 %p")
    curr_time = curr_time + " IST"
    email_body = "<html><head></head><body><p>Hello All,</p><p>Please find enclosed BCP report at " + str(curr_time) + ".<br><br>Regards,<br>Internal IT Networks</p></body></html>"
    s = smtplib.SMTP("unixmailrelay.internal.colt.net", 25)
    msg = MIMEMultipart()
    msg["From"] = "TSOperations-RPASupport@COLT.NET"
    msg["Subject"] = "BCP Hourly Report - Internet & VPN Monitoring - " + str(bcp_report_name.split("-")[1].split(".")[0]).replace("_", " ") + " @" + str(curr_time)
    msg['To'] = "Atiqur.Rahman2@COLT.NET"
    msg['Cc'] = "Atiqur.Rahman2@COLT.NET"
    msg['Bcc'] = "Atiqur.Rahman2@COLT.NET"
    # msg["To"] = "SysOpsNetwork@colt.net, CyberSecurityOperationsCentre@colt.net, IncidentManagement2@colt.net"
    # msg["Cc"] = "Ashish.Surti@colt.net, Sukant.Bhattacharjee@colt.net , Amit.Babbar@colt.net, Parvinder.Singh@colt.net, Anirudh.Kumar@colt.net, Vipul.Agarwal@colt.net, Venkatesh.Ravindran@colt.net, Manish.Kumar1@colt.net, Manohar.Singhbisht@colt.net, Harshwardhan.Deshmukh@colt.net, Ashish.Gaur@colt.net, Mohit2.Mehta@colt.net"
    # msg["Bcc"] = "Atiqur.Rahman2@COLT.NET,Sandeep.Gupta@colt.net"
    os.chdir(bcp_output_path)
    file_list = glob.glob(bcp_report_name)
    for files in file_list:
        fo = open(files, "rb")
        file_type = files.split(".")[-1]
        attach_file = email.mime.application.MIMEApplication(fo.read(), _subtype=file_type)
        fo.close()
        attach_file.add_header('Content-Disposition', 'attachment', filename=files)
        msg.attach(attach_file)
    msg.attach(MIMEText(email_body, "html"))
    s.send_message(msg,rcpt_options=['NOTIFY=NEVER'])
    s.quit()
    del msg
    gc.collect()