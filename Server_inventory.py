import boto3

import boto3
import openpyxl
import os
import csv

from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from datetime import datetime, timezone, timedelta

import smtplib

lis =[["Account","us-east-1","","","us-east-2","","","Total","",""],["","Running","Stopped","New Server this week","Running","Stopped","New Servers This week","Running","Stopped", "New Servers This week","total"]]

def loadbalancer(profile):
    last_friday = datetime.now().date() - timedelta(days=7)
    
    session = boto3.Session(profile_name=profile)
    
    profile_regions = {
        "GBS-GSD-EnterpriseApps-DevTest":["us-east-1","us-east-2"],
        "GBS-EnterpriseApps-Prod" : ["us-east-1","us-east-2"],
        "GBS-GSD-HCM-DevTest" : ["eu-west-1","eu-central-1"],
        "Step-SAP-DevTest" : ["us-east-1","us-east-2"],
        "GBS-GSD-HCM-A-Prod" : ["eu-west-1","eu-central-1"],
        "GBS-GSD-E1_A-Prod" : ["us-east-1","us-east-2"],
        "GBS-GSD-PCI-PCI-Prod":["us-east-1","us-east-2"],
        "GBS-EA-GRCHadoop-Prod":["us-east-1","us-east-2"],
        "GBS-GSD-GCR-A-DevTest":["us-east-1","us-east-2"],
        "GBS-GSDBO-Shared_APPS-A-NonPrd":["us-east-1","us-east-2"],
        "GBS-GSDBO-Shared_APPS-A-Prd":["us-east-1","us-east-2"],
        "GBS-GSD-HCM-NP-PROD":["us-east-1","us-east-2"],
        "GBS-GSD-HCM-Prod":["us-east-1","us-east-2"],
        "GBS-GSD-EA-A-FBDR":["us-east-1","us-east-2"],
        "GBS-GBS-e-sign-A-NONPRD":["us-east-1","us-east-2"],
        "GBS-GBS-e-sign-A-PRD":["us-east-1","us-east-2"]
    }
    lis2=[]
    lis2.append(profile)
    stopped_main = 0
    running_main = 0
    Last_week_main = 0
    for p in profile_regions[profile]:
        client = session.client("ec2",region_name=p)
        paginator = client.get_paginator('describe_instances')
        response_iterator = paginator.paginate()
        stopped = 0
        running = 0
        Last_week = 0
        rest = 0
        for page in response_iterator:
            if len(page["Reservations"]) != 0:
                for i in page["Reservations"]:
                    for lpty in i["Instances"]:
                        if lpty["State"]["Name"] == "running":
                            running += 1
                        elif lpty["State"]["Name"] == "stopped":
                            stopped += 1
                        else:
                            rest += 1
                        try:
                            if lpty['NetworkInterfaces'][0]['Attachment']['AttachTime'].date() > last_friday:
                                Last_week += 1
                        except:
                            print(f"excepted once in {profile}")
                        #print(lpty['LaunchTime'].date())

                        

        lis2.append(running)
        lis2.append(stopped)
        lis2.append(Last_week)
        stopped_main += stopped
        running_main += running
        Last_week_main += Last_week
    lis2.append(running_main)
    lis2.append(stopped_main)
    lis2.append(Last_week_main)
    lis2.append(running_main+stopped_main)
    lis.append(lis2)
    




    
def main():
    profiles = ["GBS-GSD-EnterpriseApps-DevTest","GBS-EnterpriseApps-Prod","Step-SAP-DevTest","GBS-GSD-PCI-PCI-Prod","GBS-GSD-EA-A-FBDR","GBS-GSD-E1_A-Prod"]
    for i in profiles:
        loadbalancer(i)

    wb = openpyxl.Workbook()

    sheet = wb.active

    le_ = len(lis)
    p = 0

    for i in lis:
        k = 0
        for j in i:
            c1 = sheet.cell(row=p+1,column=k+1)
            c1.value = str(j)
            k+=1
        p+=1

    print('success')

    wb.save("E:\crap\inventory.xlsx")

    msg = MIMEMultipart()
    msg['Subject'] = f'AWS Inventory'
    msg['From'] = 'FridayServerInventory@wolterskluwer.com'
    msg['To'] = 'venkata.khande@wolterskluwer.com, tanmay.jog@wolterskluwer.com, vc.parepelli@wolterskluwer.com'

    with open('E:\crap\inventory.xlsx', 'rb') as f:
        part = MIMEApplication(f.read())
        part.add_header('Content-Disposition', 'attachment', filename='inventory.xlsx')
        msg.attach(part)


    part = MIMEText("Hi Team, Please find the report in the attachments")
    msg.attach(part)



    # session = boto3.session.Session(profile_name='GBS-GSD-EnterpriseApps-DevTest')
    # connect = session.client('ses','us-east-1')

    try:
    #Provide the contents of the email.
        # response = connect.send_raw_email(
        #     Source='venkata.khande@wolterskluwer.com',
        #     Destinations=[
        #         'venkata.khande@wolterskluwer.com',
        #         'tanmay.jog@wolterskluwer.com'
        #     ],
        #     RawMessage={
        #         'Data':msg.as_string(),
        #     },
        # )

        s = smtplib.SMTP('smtp-nonprodrelay1.gsdwkglobal.com')

        s.send_message(msg)

       
# Display an error if something goes wrong. 
    except Exception as e:
        print(e)
    else:
        print("Email sent! Message ID:"),
        #print(response['MessageId'])
        

main()
