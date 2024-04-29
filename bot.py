#!/usr/bin/env python3

import csv; import os; import pickle; import re; import sys
from itertools import permutations; import json; import random; import time; from datetime import datetime
from datetime import timedelta; import asyncio
from typing import final; import discord
from discord import guild
import pygsheets; import traceback
from pygsheets.datarange import DataRange
import requests; from discord import File, User, Webhook
from discord.ext import commands, tasks; from discord.ext.commands import Bot, Cog, CommandNotFound, Context
from discord.utils import get; from google.oauth2 import service_account;
from googleapiclient.discovery import build; from google_auth_oauthlib.flow import InstalledAppFlow; from google.auth.transport.requests import Request
from base64 import urlsafe_b64decode, urlsafe_b64encode; from email.mime.text import MIMEText; from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage; from email.mime.audio import MIMEAudio; from email.mime.base import MIMEBase; from mimetypes import guess_type as guess_mime_type
from gcsa.google_calendar import GoogleCalendar; import gcsa.event; import gcsa.recurrence; from gcsa.calendar import CalendarListEntry; from gcsa.event import Event

sys.stdout.reconfigure(line_buffering=True)

# Request all access (permission to read/send/receive emails, manage the inbox, and more)
SCOPES = ['https://mail.google.com/',]
our_email = 'wolfie.seawolf.bot@gmail.com'

SERVICE_SCOPES = ['https://www.googleapis.com/auth/calendar','https://www.googleapis.com/auth/drive']
SERVICE_ACCOUNT_FILE = "pygsheets_key.json"

credentials = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SERVICE_SCOPES)
gc = GoogleCalendar('eumlsmdr2asjjrjm6q3vm3a5r8@group.calendar.google.com',credentials=credentials)
pto = GoogleCalendar('ec0f568262d866fb5c60e5b5b436671325a522c6eab10faa23c6b9dfb408f02d@group.calendar.google.com',credentials=credentials)
calendar = gc.get_calendar('eumlsmdr2asjjrjm6q3vm3a5r8@group.calendar.google.com')

drive_service = build('drive','v3',credentials=credentials)
on_campus_folder = "11itV36I_JvKApDKgNRWd1KPpTVRoSvvY" #UNTIL January 2025
off_campus_folder = "11mvnloCSRDxc8_3y5G3kilGxN5z0q7Ey" #UNTIL January 2025

print("Calendars:")
for calendar in gc.get_calendar_list():
    print(calendar)
print("Events:")
for event in gc:
    print(event)
print("done")

version = f'1.1.8'
signature = f'James D. Boglioli'
name = "Alpha Wolf"
Project_Maintainer = "James Boglioli (James.Boglioli@StonyBrook.edu)"
Marketing_Director = "Andrea Lebedinski (Andrea.Lebedinski@stonybrook.edu)"
Management_Emails = "James.Boglioli@stonybrook.edu, Andrea.Lebedinski@stonybrook.edu"

event_confirmation_days = 2

DISCORD_TOKEN = open("bot_token.txt","r")
DISCORD_TOKEN = DISCORD_TOKEN.read()

DEV_TOKEN = ""
try:
    DEV = open("dev_token.txt","r")
    DEV_TOKEN = DEV.read()
except:
    DEV_TOKEN = "False"


TOKEN = DISCORD_TOKEN
description = f'Slash Commands Supported - V {version} By: {signature}'

intents = discord.Intents.default(); intents.members = True
client = commands.Bot(command_prefix = ">", intents=intents)
bot = commands.Bot(command_prefix='>', description=description, intents=discord.Intents.all(), case_insensitive=True, help_command=None)

gsclient = pygsheets.authorize(service_file='pygsheets_key.json'); gsheet = gsclient.open_by_key('1n_zqs13W4IsMAAvnX12I-sFmKtS6tfTpI4_8dnym58Q')
PTOsheet = gsclient.open_by_key('1yLufq6ZppMDydmZ9ftmnuUzduqNib6Qv_p09R_3Ap2A')
api = gsclient.open_by_key("1Qi8egVV5cS5G_9WTi94A6B6ZxmThePRt54I6fXYsOGE")



contact_info = gsheet.worksheet_by_title('WOLFIE CONTACT INFO')
wolfie_schedule = gsheet.worksheet_by_title('SCHEDULE')
weekapi = api.worksheet_by_title('WeeklyEvents')
PTOlist = PTOsheet.worksheet_by_title("Formatted Data")

week_events = {}
body_dict = {"body":""}

def getOSdatetime() -> str:
    print(os.name)
    if os.name == "nt":
        print("#")
        return "#"
    else: 
        print("-")
        return "-"

monthvar = getOSdatetime()
on_ready_status = {"status":0}

@bot.event
async def on_ready(): #Has error handling
    try:
        if DEV_TOKEN == "False" and utils.AutoUpdate.is_running() == False: utils.AutoUpdate.start()
        else: print("This is a DEV environment")
        if gcal.iterate_events.is_running() == False: gcal.iterate_events.start()
        if gcal.timeoff.is_running() == False: gcal.timeoff.start()
        server = bot.get_guild(901724465476546571)
        await bot.change_presence(activity=discord.Activity(type=discord.ActivityType.watching, name=f"The Wolfie Team"))
        print("Wolfie is Watching")
        chan2 = bot.get_channel(1112127760740130876)
        if on_ready_status["status"] == 0: await chan2.send(f"Wolfie Has Restarted on Version {version}")
        on_ready_status["status"] += 1
    except Exception:
        await utils.ErrorHandler(Exception,"Startup")
    #await utils.createRemoteFolder(folderName="TEST",parentID=on_campus_folder)
    #await main()

class utils:
    async def TimeCheck(timeStart, timeEnd, day = "NA") -> bool:                # SHOULD USE FORMAT '6:00AM' OR '11:00PM' and 0-6 for Mon - Sun
        today = datetime.today().weekday()                                      # gets current day of the week
        if day != "NA":                                                         # checks if a day variable is added to the timecheck
            if today == day: dychk = True                                       # checks if the current day is the correct day to run
            else: dychk = False                                                 # what to do if it is not the correct day
        else: dychk = True                                                      # what to do if there is no daycheck
        now = datetime.now()                                                    # saves the current time
        timeNow = now.strftime('%H:%M:%S')                                      # converts the current time into readable format
        timeEnd = datetime.strptime(timeEnd, "%I:%M%p")                         # converts the end time into readable and comparable format
        timeStart = datetime.strptime(timeStart, "%I:%M%p")                     # converts the start time to readable and comparable format
        timeNow = datetime.strptime(timeNow, "%H:%M:%S")                        # converts the now time into a comparable format
        if timeStart < timeEnd:                                                 # checks if the start time is less than the end time
            return timeNow > timeStart and timeNow < timeEnd and dychk          # returns True if time is between start and end
        else:                                                                   # checks if end is less than start time (overlapping midnight)
            return dychk and (timeNow > timeStart or timeNow < timeEnd)         # returns True if time is between start and end

    async def GetEmailList() -> str:                                            # RETURNS ALL THE CURRENT WOLFIE STAFF EMAIL ADDRESSES
        x = 4                                                                   # sets starting value to begin with Maria
        value = "1"; email_list = ""                                            # sets up needed variables
        while value != "":                                                      # loops while emails are detected in list
            value = contact_info.cell(f"F{x}").value                            # gets the email from list
            if x != 5:                                                          # 
                email_list = email_list + ", " + value                          # adds new value and parenthesis
            else:                                                               # 
                email_list = value                                              # adds the first email
            x += 1                                                              # adds one to x to go down the list of emails
            value = contact_info.cell(f"F{x}").value                            # updates value again
        return email_list                                                       # returns email list to the outside function as a string

    async def wkday(date) -> str:                                               # Returns the current weekday of the day
        today = datetime.today()                                                # gets todays date
        chk = today.strftime(f"%{monthvar}m/%d/%Y").split("/")                   # turns todays date into a string and splits by /
        date = date.split("/")                                                  # splits the date by /
        if date[0] >= chk[0]:                                                   # checks if the month is greater or less than current month
            day = datetime.strptime(f"{date[0]}/{date[1]}/{chk[2]}","%m/%d/%Y") # sets date to have current year
        else:                                                                   # 
            day = datetime.strptime(f"{date[0]}/{date[1]}/{chk[2]+1}","%m/%d/%Y")#sets date tp have year after current year
        wkday = day.weekday()                                                   # gets the weekday value of the requested day
        if wkday == 0: return "Monday"                                          # here down returns correct date for day given by wkday
        elif wkday == 1: return "Tuesday"                                       #
        elif wkday == 2: return "Wednesday"                                     #
        elif wkday == 3: return "Thursday"                                      #
        elif wkday == 4: return "Friday"                                        #
        elif wkday == 5: return "Saturday"                                      #
        elif wkday == 6: return "Sunday"                                        #

    async def GetDate(date) -> datetime.date:                                   # Returns datetime object from a date in format '8/21'
        today = datetime.today()                                                # gets date of today
        chk = today.strftime(f"%{monthvar}m/%d/%Y").split("/")                  # splits string of today
        date = date.split("/")                                                  # splits date in question
        if date[0] >= chk[0]:                                                   # compares month value of both dates
            day = datetime.strptime(f"{date[0]}/{date[1]}/{chk[2]}","%m/%d/%Y") # sets year to current year
        else:                                                                   #
            day = datetime.strptime(f"{date[0]}/{date[1]}/{chk[2]+1}","%m/%d/%Y")#sets year to next year
        return day                                                              # returns date of day in question

    async def DateTimeCombine(date,time) -> datetime.date:
        mytime = datetime.strptime(time,f'%H:%M').time() #needs to have zero padding hour
        date = datetime.strptime(date,f'%m/%d/%Y') # needs to have zero padding month and day
        mydatetime = datetime.combine(date, mytime)
        return mydatetime

    async def ZeropadDatetime(typ:str,string:str) -> str:                       # Takes input as D or T
        if typ.lower() == 't':
            string = string.split(":")
            if len(string[0]) == 1:
                string[0] = "0" + string[0]
            string = string[0] + ":" + string[1]
        elif typ.lower() == 'd':
            string = string.split("/")
            if len(string[0]) == 1:
                string[0] = "0" + string[0]
            if len(string[1]) == 1:
                string[1] = "0" + string[1]
            if len(string[2]) != 4: raise TypeError("The Year is invalid")
            string = string[0] + "/" + string[1] + "/" + string[2]
        else: raise TypeError("Only D or T (date or time) are allowed")
        print(string)
        return string

    async def Convert24h(time) -> str:
        if "pm" in time.lower() and "12:" not in time:
            time = time.split(":")
            time[0] = str(int(time[0]) + 12)
            time[1] = time[1].lower().replace(" ","").replace("pm","")
        elif "am" in time.lower() and "12:" in time:
            time = time.split(":")
            time[0] = "00"
            time[1] = time[1].lower().replace(" ","").replace("am","")
        else:
            time = time.lower().replace("am","").replace("pm","").replace(" ","")
            time = time.split(":")
        time = str(time[0] + ":" + time[1])
        print(time)
        return time

    async def GetAcademicYear() -> str:
        currentMonth = datetime.now().month
        currentYear = datetime.now().year
        year = int(str(currentYear)[2:])
        if int(currentMonth) > 7: yearstr = str(year) + "-" + str(year + 1)
        else: yearstr = str(year-1) + "-" + str(year)
        return yearstr

    async def ErrorHandler(exception,task):
        # Function works as follows:
        # try: 1/0
        # except Exception: await utils.ErrorHandler(Exception,"Current Function") 
        e = traceback.format_exc(1900)
        chan = bot.get_channel(1112127760740130876)
        await chan.send(f"***Error Detected in {task}:***```{e}```")

    async def get_service(api_name, api_version, scopes, key_file_location):
        """Get a service that communicates to a Google API.

        Args:
            api_name: The name of the api to connect to.
            api_version: The api version to connect to.
            scopes: A list auth scopes to authorize for the application.
            key_file_location: The path to a valid service account JSON key file.

        Returns:
            A service that is connected to the specified API.
        """
        credentials = service_account.Credentials.from_service_account_file(
        key_file_location)
        scoped_credentials = credentials.with_scopes(scopes)
        # Build the service object.
        service = build(api_name, api_version, credentials=scoped_credentials)
        return service

    async def createRemoteFolder(folderName, parentID = None):
        # Create a folder on Drive, returns the newely created folders ID
        body = {
          'name': folderName,
          'mimeType': "application/vnd.google-apps.folder"
        }
        if parentID:
            body['parents'] = [parentID]
        root_folder = drive_service.files().create(body = body).execute()
        return root_folder['id']

    async def createEventFolder(event_name,event_date,event_type="on_campus/off_campus",spotter=""):
        folder_name = f"{event_date} - {event_name}"
        if spotter != "": 
            spotter = spotter.replace("\n", " ").split(" ")
            spotter = list(filter(lambda x: len(x) > 0, spotter))
            xp = round(len(spotter)/2)
            xpp = 2
            spotter_name = spotter[0]
            while xpp <= xp:
                spotter_name = spotter_name + ", " + spotter[xpp]
                spotter_name += 2
            folder_name = folder_name + f" [{spotter_name}]"
        if event_type == "on_campus": folder = on_campus_folder
        if event_type == "off_campus": folder = off_campus_folder
        root = await utils.createRemoteFolder(folder_name,folder)
        photoFolder = await utils.createRemoteFolder("Photos",root)


    @tasks.loop(minutes=60)
    async def AutoUpdate() -> bool:
        timecheck1 = await utils.TimeCheck('3:00am','3:15am')
        timecheck2 = await utils.TimeCheck('4:00am','4:15am')
        if timecheck1 == False and timecheck2 == False:
            print("Looking For Updates...")
            nl = '\n'
            qint = random.randint(1,2000)
            update_url = f"https://raw.githubusercontent.com/OhNo2/NewWolfieBot/main/bot.py?v={qint}"
            version_url = f"https://raw.githubusercontent.com/OhNo2/NewWolfieBot/main/version.txt?v={qint}"
            r = requests.get(update_url,allow_redirects=True)
            await asyncio.sleep(0)
            v = requests.get(version_url,allow_redirects=True)
            await asyncio.sleep(0)
            ver = v.text.replace('"',"").replace(nl,"").split(" = ")[1]
            if ver != version:
                print("Update Available, Downloading...")
                chan2 = bot.get_channel(1112127760740130876)
                await chan2.send(f"Update Found! Automatically Updating to Version {ver}")
                await asyncio.sleep(0)
                open('bot.py','wb').write(r.content)
                await asyncio.sleep(0)
                sys.exit()
                #try:
                #    os.execv(__file__,sys.argv)
                #except:
                #    os.execv(sys.executable, ['python'] + sys.argv)
            else:
                print("Alpha Wolf is Up to Date!")

class gcal:
    async def create_event(title:str,date:str,start_time:str,end_time:str,signups:str,description:str,end_date:str="NA",calSelect="gc"):
        if end_date == "NA":
            if start_time.lower() != "all day":
                start_time = await utils.ZeropadDatetime("T",start_time)
                end_time = await utils.ZeropadDatetime("T",end_time)
                date = await utils.ZeropadDatetime("D",date)
                start = await utils.DateTimeCombine(date,start_time)
                end = await utils.DateTimeCombine(date,end_time)
            else:
                date = await utils.ZeropadDatetime("D",date)
                start = datetime.strptime(date,f'%m/%d/%Y')
                end = datetime.strptime(date,f'%m/%d/%Y')
        else:
            if start_time.lower() != "all day":
                start_time = await utils.ZeropadDatetime("T",start_time)
                end_time = await utils.ZeropadDatetime("T",end_time)
                start_date = await utils.ZeropadDatetime("D",date)
                end_date = await utils.ZeropadDatetime("D",end_date)
                start = await utils.DateTimeCombine(start_date,start_time)
                end = await utils.DateTimeCombine(end_date,end_time)
            else:
                start_date = await utils.ZeropadDatetime("D",date)
                end_date = await utils.ZeropadDatetime("D",end_date)
                start = datetime.strptime(start_date,f'%m/%d/%Y')
                end = datetime.strptime(end_date,f'%m/%d/%Y')
        event = Event(title,start,end,location=signups,description=description)
        if calSelect == "gc": gc.add_event(event)
        if calSelect == "pto": pto.add_event(event)
        else: "ERROR: INVALID CALENDAR"
        print(f"Event titled {title} added to calendar")

    async def update_event(event,row:int,title:str,date:str,start_time:str,end_time:str,signups:str,description:str):
        pass

    async def delete_event():
        pass

    @tasks.loop(minutes=15)
    async def iterate_events():
        try:
            print("Starting Search...")
            nl = '\n'
            today = datetime.now().strftime("%m/%d/%Y")
            timecheck = await utils.TimeCheck('12:00am','12:15am')
            #timecheck = True
            if timecheck == True:
                # Get list of current events to check events against
                #events = gc.get_events(datetime.today(), datetime.today() + timedelta(days=180))
                #event_lst = list(events)
                #str_event_lst = []
                #for event in event_lst:
                #    event = str(event).split(" - ")
                #    sublst = event[0].split(" ")
                #    sublst[1] = sublst[1].split("-")[0]
                #    sublst[0] = sublst[0].split("-")
                #    sublst[0] = str(sublst[0][1] + "/" + sublst[0][2] + "/" + sublst[0][0])
                #    sublst[0] = await utils.ZeropadDatetime("D",sublst[0])
                #    sublst[1] = sublst[1].split(":"); sublst[1] = str(sublst[1][0]) + ":" + str(sublst[1][1])
                #    l = [event[1],sublst[0],sublst[1]]
                #    print(l)
                #    str_event_lst.append(l)
                #print(events)
                #print(event_lst)
                #print(str_event_lst)
                # Begin checking the spreadsheet for current events
                x = 0; y = True
                unf_evt = discord.Embed(title="Unfilled Events Eligible for Assigmnet",description=datetime.now().strftime("%m/%d/%Y"),url="https://docs.google.com/spreadsheets/d/1n_zqs13W4IsMAAvnX12I-sFmKtS6tfTpI4_8dnym58Q/edit?usp=sharing")
                wk_unf_evts = ""
                wk_unf = 1
                unf = False
                while y == True:
                    pause = 3
                    x += 1
                    print(x)
                    date = str(wolfie_schedule.cell(f"B{x}").value)
                    date = date.replace(" ","")
                    try:
                        mydate = await utils.ZeropadDatetime("D",str(date))
                        dtdate = datetime.strptime(mydate,"%m/%d/%Y")
                    except:
                        dtdate = datetime.strptime("04/10/2002", "%m/%d/%Y")
                        print("Not A Date")
                    if datetime.strptime(today,"%m/%d/%Y") <= dtdate: # Works if the date of the event is after today
                        title = wolfie_schedule.cell(f"C{x}").value
                        location = wolfie_schedule.cell(f"D{x}").value
                        stime = wolfie_schedule.cell(f"E{x}").value
                        start_time = str(await utils.Convert24h(str(stime)))
                        etime = wolfie_schedule.cell(f"F{x}").value
                        end_time = str(await utils.Convert24h(str(etime)))
                        wolfie = wolfie_schedule.cell(f"I{x}").value
                        spotter = wolfie_schedule.cell(f"J{x}").value
                        requestor = wolfie_schedule.cell(f"K{x}").value
                        confirmed = wolfie_schedule.cell(f"M{x}").value
                        additional_info = wolfie_schedule.cell(f"N{x}").value
                        cal_created = wolfie_schedule.cell(f"P{x}").value
                        signups = wolfie + " " + spotter
                        signups = signups.replace(nl," ").split(" ")
                        signups  = [i for i in signups if i]
                        z = ""
                        xx = len(signups)
                        xxx = 0
                        while xxx < xx:
                            if xxx%2 == 0: z = z + signups[xxx] + " "
                            elif xxx != xx - 1: z = z + signups[xxx] + ", "
                            else: z = z + signups[xxx]
                            xxx += 1
                        signups = z
                        print(signups)
                        if confirmed == "--": confirm = "Yes"
                        elif confirmed == "": confirm = "No"
                        elif confirmed == "x": confirm = "No"
                        elif "/" in confirmed: confirm = "Yes"
                        else: confirm = "No" 
                        description = f"Location: {location}{nl}{nl}Requestor: {requestor}{nl}{nl}Event is Confirmed: {confirm}{nl}{nl}Additional Notes: {additional_info}"
                        #Begin to handle the events
                        if cal_created.lower() != "x": # The event has not been created yet
                            pause = 11
                            await gcal.create_event(title,date,start_time,end_time,signups, description)
                            embed = discord.Embed(title=title,description=f'Location: {location}',url="https://docs.google.com/spreadsheets/d/1n_zqs13W4IsMAAvnX12I-sFmKtS6tfTpI4_8dnym58Q/edit?usp=sharing")
                            embed.add_field(name="Event Date:",value=date)
                            embed.add_field(name="Event Duration:",value=f'{start_time}-{end_time}')
                            embed.add_field(name="Requestor Contact:",value=requestor,inline=False)
                            if additional_info != "": embed.add_field(name="Additional Info:",value=additional_info,inline=False)
                            embed.set_footer(text="Info subject to change. Acts as event creation reciept. Check spreadsheet for accurate info")
                            sheetchan = bot.get_channel(902627884995321937)
                            await sheetchan.send(embed=embed)
                            wolfie_schedule.update_value(f"P{x}","X")
                            wolfie_schedule.update_value(f"Q{x}",f"{date}")
                            wolfie_schedule.update_value(f"R{x}",start_time)
                            wolfie_schedule.update_value(f"S{x}",wolfie)
                            wolfie_schedule.update_value(f"T{x}",spotter)
                            wolfie_schedule.update_value(f"U{x}",additional_info)
                            wolfie_schedule.update_value(f"V{x}",confirmed)
                        elif wolfie_schedule.cell(f"A{x}").value != "": #checks event that has already been created
                            pause = 19
                            print(f"'{title}'")
                            print(f"'{date}'")
                            olddate = wolfie_schedule.cell(f"Q{x}").value
                            date = date.replace(" ","")
                            try:
                                myolddate = await utils.ZeropadDatetime("D",str(date))
                                dtolddate = datetime.strptime(myolddate,"%m/%d/%Y")
                            except:
                                dtolddate = datetime.strptime("04/10/2002", "%m/%d/%Y")
                            editevt = gc.get_events(dtolddate - timedelta(days=1),dtolddate + timedelta(days=1),query=title,timezone="America/New_York")
                            m = 0
                            for event in editevt:
                                edevt = event
                                m += 1
                            if m == 0:
                                olddate = olddate.replace(" ","")
                                editevt = gc.get_events(dtolddate - timedelta(days=1),dtolddate + timedelta(days=1),query=title,timezone="America/New_York")
                                for event in editevt:
                                    edevt = event
                            try:        
                                print(edevt)
                            except:
                                wolfie_schedule.update_value(f"P{x}","")
                                pause += 1
                                continue
                            datechk = bool(wolfie_schedule.cell(f"Q{x}").value == date)
                            startchk = bool(wolfie_schedule.cell(f"R{x}").value == start_time)
                            wolfchk = bool(wolfie_schedule.cell(f"S{x}").value == wolfie)
                            spotchk = bool(wolfie_schedule.cell(f"T{x}").value == spotter)
                            addchk = bool(wolfie_schedule.cell(f"U{x}").value == additional_info)
                            confchk = bool(wolfie_schedule.cell(f"V{x}").value == confirmed)
                            chklst = [datechk, startchk, wolfchk, spotchk, addchk, confchk]
                            z = 0
                            for chk in chklst:
                                if chk == False:
                                    print(f"check {chk}")
                                    edited = False
                                    if z == 0 or z == 1:
                                        edevt.start = await utils.DateTimeCombine(date,start_time)
                                        edevt.end = await utils.DateTimeCombine(date,end_time)
                                        wolfie_schedule.update_value(f"Q{x}",f"{date}")
                                        wolfie_schedule.update_value(f"R{x}",start_time)
                                        edited = True
                                    elif z == 2 or z == 3:
                                        edevt.location = signups
                                        wolfie_schedule.update_value(f"S{x}",wolfie)
                                        wolfie_schedule.update_value(f"T{x}",spotter)
                                        edited = True
                                    elif z == 4 or z == 5:
                                        edevt.description = f"Location: {location}{nl}{nl}Requestor: {requestor}{nl}{nl}Event is Confirmed: {confirm}{nl}{nl}Additional Notes: {additional_info}"
                                        wolfie_schedule.update_value(f"U{x}",additional_info)
                                        wolfie_schedule.update_value(f"V{x}",confirmed)
                                        edited = True
                                    gc.update_event(edevt)
                                    print("Event Edited")
                                    if edited == True:
                                        editevt = gc.get_events(datetime.strptime(date, "%m/%d/%Y"),datetime.strptime(date, "%m/%d/%Y") + timedelta(days=1),query=title,timezone="America/New_York")
                                        m = 0
                                        for event in editevt:
                                            edevt = event
                                            m += 1
                                        if m == 0:
                                            olddate = olddate.replace(" ","")
                                            editevt = gc.get_events(datetime.strptime(olddate, "%m/%d/%Y"),datetime.strptime(olddate, "%m/%d/%Y") + timedelta(days=1),query=title,timezone="America/New_York")
                                            for event in editevt:
                                                edevt = event
                                        print(edevt)
                                z += 1
                            if dtdate <= datetime.now() + timedelta(days=7):
                                desc = f"{date}: {start_time}-{end_time}"
                                if wolfie == "" and spotter == "": unf_evt.add_field(name=f'{title} - W & S Required',value=desc,inline=False)
                                elif wolfie == "": unf_evt.add_field(name=f'{title} - Wolfie Required',value=desc,inline=False)
                                elif spotter == "": unf_evt.add_field(name=f'{title} - Spotter Required',value=desc,inline=False)
                                if wolfie == "" or spotter == "": unf = True
                            weekday = datetime.today().weekday()
                            if  weekday == 6 and dtdate <= datetime.now() + timedelta(days=15) and (wolfie == "" or spotter == ""):
                                wkday = dtdate.weekday()
                                if wkday == 0: wkday = "Mo"
                                elif wkday == 1: wkday = "Tu"
                                elif wkday == 2: wkday = "We"
                                elif wkday == 3: wkday = "Th"
                                elif wkday == 4: wkday = "Fr"
                                elif wkday == 5: wkday = "Sa"
                                elif wkday == 6: wkday = "Su"
                                truncdat = dtdate.strftime("%m/%d")
                                sstime = stime.lower().replace(" ","")
                                eetime = etime.lower().replace(" ","")
                                if len(sstime) < 7: sstime = "0" + sstime
                                if len(eetime) < 7: eetime = "0" + eetime
                                evt = f"{wk_unf}. {wkday} {truncdat} @ {sstime} - {eetime} "
                                if wolfie == "" and spotter == "":evt = evt + "(1x W, 1x S)"
                                elif wolfie == "": evt = evt + "(1x W)"
                                elif spotter == "": evt = evt + "(1x S)"
                                evt = evt + nl
                                wk_unf_evts = wk_unf_evts + evt
                                wk_unf += 1
                        if datetime.strptime(today,"%m/%d/%Y") + timedelta(days=1) == dtdate:
                            description = additional_info.lower()
                            eventType = wolfie_schedule.cell(f"G{x}").value.lower()
                            if "off campus" in description or "off-campus" in description: evtType = "off_campus"
                            elif "bb" in eventType or "football" in eventType or "mascot madness" in eventType or "meeting" in eventType or "tournament" in eventType or "uca" in eventType or "rehersal" in eventType or "practice" in eventType: evtType = "none"
                            else: evtType = "on_campus"
                            event_date = datetime.strftime(dtdate,"%Y/%m/%d")
                            if evtType != "none": await utils.createEventFolder(title,event_date,evtType,spotter)
                    elif dtdate != datetime.strptime("04/10/2002", "%m/%d/%Y") and datetime.strptime(today,"%m/%d/%Y") > dtdate: # Works if the event has already happened
                        pause = 2
                        sheet = await utils.GetAcademicYear()
                        sheet = gsheet.worksheet_by_title(f'COMPLETED {sheet} EVENTS')
                        shrow = int(sheet.cell("AD1").value)
                        columns = ["a","b","c","d","e","f","g","i","j","k","l","m","n","o"]
                        pause += 2
                        for col in columns:
                            if col == "k":
                                requester = wolfie_schedule.cell(f"{col}{x}").value
                                await asyncio.sleep(1)
                            if col =="c":
                                event_name = wolfie_schedule.cell(f"{col}{x}").value
                                await asyncio.sleep(1)
                            sheet.update_value(f"{col}{shrow}",wolfie_schedule.cell(f"{col}{x}").value)
                            await asyncio.sleep(1)
                        if "@" in requester:
                            subject = f"Wolfie Satisfaction Form - {event_name}"
                            requester_name = requester.split("@")[0].split(".")[0].capitalize()
                            subject = subject.replace(" ","%20")
                            email_date = dtdate.strftime("%Y-%m-%d")
                            nl = '\n'
                            body = f"Hello {requester_name},%0D%0A%0D%0AWolfie had a lot of fun at {event_name}! If you have any events in the future that you would like Wolfie to attend, please put in a request (link in my signature).%0D%0A%0D%0AWe are always looking to improve how Wolfie does at events, so I was hoping you could fill out a brief survey on how Wolfie did! Any comments help us to make events better in the future! Please take the time to fill out this survey:%0D%0A%0D%0Ahttps://docs.google.com/forms/d/e/1FAIpQLSe2C0f17HLMSU1fEiuK93bUS65nh8UYuJ46YllXC4gnXpYzSA/viewform?usp=pp_url&entry.708486352={event_name.replace(' ','%20')}&entry.1669706537={email_date}"
                            body = body.replace("%0D%0A",nl)
                            email_link = f"{body}"
                            email_embed = discord.Embed(title=f"Wolfie Satisfaction Form Email - {event_name}",description=f"send to: {requester}")
                            email_chan = bot.get_channel(1137056916032475186)
                            await email_chan.send(embed=email_embed)
                            await email_chan.send(email_link)
                        wolfie_schedule.delete_rows(x)
                        x -= 1
                        #wolfie_schedule.update_dimensions_visibility(x,x,dimension='ROWS',hidden=True)
                        print("Moved Row")
                    elif wolfie_schedule.cell(f"A{x}").value == "" and wolfie_schedule.cell(f"A{x+1}").value == "": # figures out that the bot has reached the end of the dates to process
                        pause = 3
                        y = False
                    print(f"{pause}s pause")
                    await asyncio.sleep(pause)
                chan = bot.get_channel(902627543864205342)
                chan2 = bot.get_channel(1074749682414272652)
                if unf == True: await chan2.send(embed=unf_evt)
                if wk_unf > 1:
                    wk_unf_embd = discord.Embed(title="Unfilled Events In The Coming Two Weeks",description=f"```{wk_unf_evts}```",url="https://docs.google.com/spreadsheets/d/1n_zqs13W4IsMAAvnX12I-sFmKtS6tfTpI4_8dnym58Q/edit?usp=sharing")
                    await chan2.send(embed=wk_unf_embd)
                await chan.send("Event calendar is now up-to-date!")
                print("Search Completed")
            else:
                print("Timecheck is False")
        except Exception: await utils.ErrorHandler(Exception,"Main Loop (iterate_events)")

    @tasks.loop(minutes=15)
    async def timeoff():
        try:
            print("Checking Time Off")
            nl = '\n'
            chan = bot.get_channel(1074749595931906100)
            today = datetime.now().strftime("%m/%d/%Y")
            timecheck = await utils.TimeCheck('4:00am','4:15am')
            #timecheck = True
            if timecheck == True:
                x = 1; y = True
                x= int(PTOlist.cell(f"J1").value)
                while y == True:
                    pause = 2
                    x += 1
                    print(x)
                    name = str(PTOlist.cell(f"A{x}").value)
                    evtchk = str(PTOlist.cell(f"G{x}").value)
                    if name == "": break
                    elif evtchk != "": pass
                    elif evtchk == "": # Runs if there is no event created and a name
                        pause += 5
                        start_date = str(PTOlist.cell(f"B{x}").value)
                        end_date = str(PTOlist.cell(f"D{x}").value)
                        description = str(PTOlist.cell(f"F{x}").value)
                        try:
                            start_time = await utils.Convert24h(str(PTOlist.cell(f"C{x}").value))
                        except:
                            start_time = "All Day"
                        try:
                            end_time = await utils.Convert24h(str(PTOlist.cell(f"E{x}").value))
                        except:
                            end_time = "All Day"
                        await gcal.create_event(f"Time Off - {name}",start_date,start_time,end_time,"",description,end_date,"pto")
                        print(f"{name} has time-off added to calendar")
                        PTOlist.update_value(f"G{x}","X")
                        embed=discord.Embed(title=f'Time Off Request: {name}', color=0x00ff00)
                        embed.add_field(name="Start Date/Time:", value=f"{start_date} - {start_time}", inline=False)
                        embed.add_field(name="End Date/Time:", value=f"{end_date} - {end_time}", inline=False)
                        embed.add_field(name="Reason For Time Off:",value=description,inline=False)
                        await chan.send(embed=embed)
                    print(f"{pause}s pause")
                    await asyncio.sleep(pause)
                print("All New PTO Updated")
                chan2 = bot.get_channel(902627543864205342)
                await chan2.send("PTO calendar is now up-to-date!")
            else: print("Timecheck False")
        except Exception: await utils.ErrorHandler(Exception,"Time Off Calendar")

class management_utils:
    @tasks.loop(minutes=15)
    async def weekly_email():
        pass

    @tasks.loop(minutes=15)
    async def unfilled_events(wolfie="",spotter="",lst=[],finished=False):
        pass


class discord_cmds:
    @bot.command(pass_context=True)
    async def hyperlink(ctx, link, text):
        embed=discord.Embed(title=text,url=link)
        await ctx.send(embed=embed)

async def main():
    print("This is for testing only! Beginning Testing...")
    #await WolfieAutomation.LookForNewEmails()
    #await WolfieAutomation.WeeklyEventEmail()
    #await WolfieAutomation.EmailConfirmation()
    #await utils.GetEmailList()
    #await WolfieAutomation.DailyCheck()
    #await WolfieAutomation.Check4Requests()
    #await gcal.create_event("test","11/10/2022","12:00","14:00","James Boglioli, Maddie Smyth","test")
    #await utils.GetAcademicYear()
    #await utils.AutoUpdate()
    print("Testing has finished")

#try:
#    service = gmail.authenticate()
#except:
#    print("Failed to authenticate!")


bot.run(TOKEN)

# What the bot needs to do:
#
# (3) Use responded emails to confirm the event with the requester
# (5) Accept WOLFIE APPLICATIONS and begin preliminary processing from inbox
# (6) Allow TEXT confirmation and signup
# (8) Give users signed up for the same event contact info for eachother in the confirmation email
#
# (8) Allow event confirmation/notification/etc to be done via Discord
#
# Email title for confirmation emails: f"CONFIRM WOLFIE SHIFT -- {shift_date} -- {event_name}"
# Email body for confirmmation emails: f"Event: {event_title}{nl}Event Location: {event_location}{nl}Event Time: {event_time}{nl}{nl}Assigned Wolfie: {a_wolfie}{nl}Assigned Spotter: {a_spotter}{nl}{nl}Please respond 'YES' to this email to confirm your shift, or 'NO' to cancel"
#
# Email title for weekly emails: f"WOLFIE EVENTS -- {monday_date} - {sunday_date}"
# 


# WHAT THE BOT CURRENTLY DOES:
# (1) Sends out emails every week (Friday) listing every event, and asking people to sign up
#       - Uses the wolfie contact information to gather email list
#       - Allows responses to email that signs the user up for a slot on the event
#       - Sends confirmation email to signup email giving user basic info about the shift
#           - Puts name of respondant into the spreadsheet
#
# (2) Sends out confirmation emails for shifts 3 days prior to the shift to confirm availability and remind staff
#       - Accepts response emails confirming availability
#           - Fills out confirmation on the spreadsheet (for potential use emailing requestor for confirmation)
#           - Gives a list of all the names of people working the shift wih you, so you can communicate about the shift
#
# (3) Processes appearance request forms
#       - Fills out spreadsheet with formatted information
#       - Emails management email list confirming processing
#       - sends out a new-shift alert for shifts that are before the next Friday email
#           - allows responses to sign up for the shift
#       - detects if a shift has too quick of a turnaround to fill
#           - alerts management
#           - does not add the event to the spreadsheet

#OLD CLASSES
class gmail:
    def authenticate():
        creds = None
        # the file token.pickle stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first time
        if os.path.exists("token.pickle"):
            with open("token.pickle", "rb") as token:
                creds = pickle.load(token)
        # if there are no (valid) credentials availablle, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)
            # save the credentials for the next run
            with open("token.pickle", "wb") as token:
                pickle.dump(creds, token)
        return build('gmail', 'v1', credentials=creds)

    def add_attachment(message, filename):
        content_type, encoding = guess_mime_type(filename)
        if content_type is None or encoding is not None:
            content_type = 'application/octet-stream'
        main_type, sub_type = content_type.split('/', 1)
        if main_type == 'text':
            fp = open(filename, 'rb')
            msg = MIMEText(fp.read().decode(), _subtype=sub_type)
            fp.close()
        elif main_type == 'image':
            fp = open(filename, 'rb')
            msg = MIMEImage(fp.read(), _subtype=sub_type)
            fp.close()
        elif main_type == 'audio':
            fp = open(filename, 'rb')
            msg = MIMEAudio(fp.read(), _subtype=sub_type)
            fp.close()
        else:
            fp = open(filename, 'rb')
            msg = MIMEBase(main_type, sub_type)
            msg.set_payload(fp.read())
            fp.close()
        filename = os.path.basename(filename)
        msg.add_header('Content-Disposition', 'attachment', filename=filename)
        message.attach(msg)

    def send(service, destination, obj, body, attachments=[]):
        return service.users().messages().send(
          userId="me",
          body=gmail.build_message(destination, obj, body, attachments)
        ).execute()

    def build_message(destination, obj, body, attachments=[]):
        if not attachments: # no attachments given
            message = MIMEText(body)
            message['bcc'] = destination
            message['from'] = our_email
            message['subject'] = obj
        else:
            message = MIMEMultipart()
            message['bcc'] = destination
            message['from'] = our_email
            message['subject'] = obj
            message.attach(MIMEText(body))
            for filename in attachments:
                gmail.add_attachment(message, filename)
        return {'raw': urlsafe_b64encode(message.as_bytes()).decode()}

    def search_messages(service, query, tags = ""):
        if tags != "": query = tags + " " + query
        result = service.users().messages().list(userId='me',q=query).execute()
        messages = [ ]
        if 'messages' in result:
            messages.extend(result['messages'])
        while 'nextPageToken' in result:
            page_token = result['nextPageToken']
            result = service.users().messages().list(userId='me',q=query, pageToken=page_token).execute()
            if 'messages' in result:
                messages.extend(result['messages'])
        return messages

    def get_size_format(b, factor=1024, suffix="B"):
        """
        Scale bytes to its proper byte format
        e.g:
            1253656 => '1.20MB'
            1253656678 => '1.17GB'
        """
        for unit in ["", "K", "M", "G", "T", "P", "E", "Z"]:
            if b < factor:
                return f"{b:.2f}{unit}{suffix}"
            b /= factor
        return f"{b:.2f}Y{suffix}"

    def clean(text):
        # clean text for creating a folder
        return "".join(c if c.isalnum() else "_" for c in text)

    def parse_parts(service, parts, folder_name, message):
        """
        Utility function that parses the content of an email partition
        """
        if parts:
            for part in parts:
                filename = part.get("filename")
                mimeType = part.get("mimeType")
                body = part.get("body")
                data = body.get("data")
                file_size = body.get("size")
                part_headers = part.get("headers")
                if part.get("parts"):
                    # recursively call this function when we see that a part
                    # has parts inside
                    gmail.parse_parts(service, part.get("parts"), folder_name, message)
                if mimeType == "text/plain":
                    # if the email part is text plain
                    if data:
                        text = urlsafe_b64decode(data).decode()
                        print(text)
                        body_dict["body"] = text
                        return text
                elif mimeType == "text/html":
                    # if the email part is an HTML content
                    # save the HTML file and optionally open it in the browser
                    if not filename:
                        filename = "index.html"
                    filepath = os.path.join(folder_name, filename)
                    print("Saving HTML to", filepath)
                else:
                    # attachment other than a plain text or HTML
                    for part_header in part_headers:
                        part_header_name = part_header.get("name")
                        part_header_value = part_header.get("value")
                        if part_header_name == "Content-Disposition":
                            if "attachment" in part_header_value:
                                # we get the attachment ID 
                                # and make another request to get the attachment itself
                                print("Saving the file:", filename, "size:", gmail.get_size_format(file_size))
                                attachment_id = body.get("attachmentId")
                                attachment = service.users().messages() \
                                            .attachments().get(id=attachment_id, userId='me', messageId=message['id']).execute()
                                data = attachment.get("data")
                                filepath = os.path.join(folder_name, filename)
                                if data:
                                    pass

    def read_message(service, message):
        """
        This function takes Gmail API `service` and the given `message_id` and does the following:
            - Downloads the content of the email
            - Prints email basic information (To, From, Subject & Date) and plain/text parts
            - Creates a folder for each email based on the subject
            - Downloads text/html content (if available) and saves it under the folder created as index.html
            - Downloads any file that is attached to the email and saves it in the folder created
        """
        msg = service.users().messages().get(userId='me', id=message['id'], format='full').execute()
        # parts can be the message body, or attachments
        payload = msg['payload']
        headers = payload.get("headers")
        parts = payload.get("parts")
        folder_name = "email"
        has_subject = False
        l = {"to":"","from":"","date":"","subject":"","body":""}
        if headers:
            # this section prints email basic info & creates a folder for the email
            for header in headers:
                name = header.get("name")
                value = header.get("value")
                if name.lower() == 'from':
                    # we print the From address
                    print("From:", value)
                    l["from"] = value
                if name.lower() == "to":
                    # we print the To address
                    print("To:", value)
                    l["to"] = value
                if name.lower() == "subject":
                    # make our boolean True, the email has "subject"
                    has_subject = True
                    # make a directory with the name of the subject
                    folder_name = gmail.clean(value)
                    print("Subject:", value)
                    l["subject"] = value
                if name.lower() == "date":
                    # we print the date when the message was sent
                    print("Date:", value)
                    l["date"] = value
        if not has_subject:
            # if the email does not have a subject, then make a folder with "email" name
            # since folders are created based on subjects
            if not os.path.isdir(folder_name):
                os.mkdir(folder_name)
        value = gmail.parse_parts(service, parts, folder_name, message)
        l["body"] = value
        if l["body"] == None:
            l["body"] = body_dict["body"]
        print("="*50)
        print(l)
        return l

    def mark_as_read(service, query):
        messages_to_mark = gmail.search_messages(service, query, "in:unread")
        return service.users().messages().batchModify(
        userId='me',
        body={
              'ids': [ msg['id'] for msg in messages_to_mark ],
              'removeLabelIds': ['UNREAD']
        }
        ).execute()

    def mark_as_unread(service, query):
        messages_to_mark = gmail.search_messages(service, query)
        # add the label UNREAD to each of the search results
        return service.users().messages().batchModify(
            userId='me',
            body={
                'ids': [ msg['id'] for msg in messages_to_mark ],
                'addLabelIds': ['UNREAD']
            }
        ).execute()

    def delete_messages(service, query):
        messages_to_delete = gmail.search_messages(service, query)
        # it's possible to delete a single message with the delete API, like this:
        # service.users().messages().delete(userId='me', id=msg['id'])
        # but it's also possible to delete all the selected messages with one query, batchDelete
        return service.users().messages().batchDelete(
        userId='me',
        body={
              'ids': [ msg['id'] for msg in messages_to_delete]
        }
        ).execute()
class WolfieAutomation:
    #@tasks.loop(minutes=2)
    async def LookForNewEmails():                                               # checks to see if there are any unread emails
        try:
            m = gmail.search_messages(service,"","in:unread")
        except:
            m = []
        if len(m) > 0:
            await WolfieAutomation.Check4Requests()
            await WolfieAutomation.EmailConfirmation()
            print("EMAIL SEARCH FINISHED")
            try: gmail.mark_as_read(service, "in:unread")
            except: pass

    #@tasks.loop(minutes=60)
    async def WeeklyEventEmail():                                               # Sends out email Friday Morning listing all the shifts for the next week
        # check sheet for all events that are occuring between monday and sunday
        # send out a list of weekly events that you can sign up for
        # include empty and full positions listed
        nl = '\n'
        timecheck = await utils.TimeCheck('8:00am','9:00am',4)
        if timecheck == True:
            week_events.clear()
            weekapi.delete_cols(2,4)
            weekapi.insert_cols(2,4)
            today = datetime.today() + timedelta(days=3)
            startdate = today.strftime(f"%{monthvar}m/%d")
            sunday = today + timedelta(days=6)
            enddate = sunday.strftime(f"%{monthvar}m/%d")
            print(f"{startdate} || {enddate}")
            weekapi.update_value('A2',startdate)
            weekapi.update_value('A4',enddate)
            day = startdate; dayx = today
            x = 0
            dates = []
            while x < 7:
                print(day)
                dates = dates + pygsheets.Worksheet.find(wolfie_schedule,f'{day}',cols = (1,2),matchEntireCell=True)
                dayx = dayx + timedelta(days=1); day = dayx.strftime(f"%{monthvar}m/%d")
                x += 1
            x = 1
            for cell in dates:
                slots = 0
                date = cell.value
                row = cell.row
                time = wolfie_schedule.cell(f"E{row}").value
                location = wolfie_schedule.cell(f"D{row}").value
                name = wolfie_schedule.cell(f"C{row}").value
                weekapi.update_value(f'B{x}',date)
                weekapi.update_value(f'C{x}',time)
                weekapi.update_value(f'D{x}',location)
                weekapi.update_value(f'E{x}',name)
                if wolfie_schedule.cell(f"I{row}").value == "": slots += 1
                wolfies = wolfie_schedule.cell(f"H{row}").value
                if wolfies == "": slots += int(wolfie_schedule.cell(f"G{row}").value)
                else:
                    wolfies = wolfies.split(",")
                    slots += int(wolfie_schedule.cell(f"G{row}").value) - len(wolfies)
                week_events[x] = [slots, date, time, name, location]
                x += 1
            print(week_events)
            body = ""
            y = week_events
            z = 1
            while z < x:
                body = body + nl + f"{z}: ({y[z][0]}ppl) - {y[z][1]} {y[z][2]} - {y[z][3]} [{y[z][4]}"[:60] + "]"
                z += 1
            body = f'The Following events need to be filled this week:{nl}{nl}' + body + nl + nl + f'Respond to this email with the event number you wish to take{nl}(EMAIL SHOULD ONLY CONTAIN EVENT NUMBER)'
            print(body)
            gmail.send(service, f"{await utils.GetEmailList()}" ,f"WOLFIE EVENTS -- {startdate} - {enddate}",body)

    #@tasks.loop(minutes=60)
    async def DailyCheck():                                                     # Sends out email asking you to confirm your shift availability
        # check for events on calendar day (2 days in advance)
        # send out emails to check for confirmation of the signed up users
        nl = "\n"
        timecheck = await utils.TimeCheck('8:30am','9:30am')
        if timecheck == True:
            today = datetime.today() + timedelta(days=event_confirmation_days)
            eventdate = today.strftime(f"#{monthvar}m/%d")
            dates = pygsheets.Worksheet.find(wolfie_schedule,f'{eventdate}',cols = (1,2),matchEntireCell=True)
            x = 5
            value = "1"
            email_list = {}
            while value != "":
                value = contact_info.cell(f"F{x}").value
                worker_name = contact_info.cell(f"A{x}").value.replace(" ","")
                email_list[worker_name.lower()] = value.lower()
                x += 1
                value = contact_info.cell(f"F{x}").value
            print(email_list)
            for x in dates:
                emails = []
                row = x.row
                wolfies = wolfie_schedule.cell(f"H{row}").value; spotters = wolfie_schedule.cell(f"I{row}").value
                event_name = wolfie_schedule.cell(f"C{row}").value
                names = wolfies + ", " + spotters
                names = names.replace(" ","").split(",")
                for name in names:
                    try:
                        email = email_list[name.lower()]
                        print(email)
                        emails.append(email)
                    except KeyError:
                        print("Not in Contact Info")
                    except:
                        print("Position not filled")
                emails = str(emails).replace("[","").replace("]","").replace("'","").replace('"',"")
                body = f"To confirm your shift, reply to this email! (Any response works) {nl}{nl}All the people working your shift are: {names}{nl}{nl}More information about your shift can be found on the Wolfie scheduling spreadsheet. {nl}{nl}If you have a change in availability or can't work this shift, please reach out to {Marketing_Director} and {Project_Maintainer} ASAP!"
                gmail.send(service, emails, f"CONFIRM WOLFIE SHIFT -- {eventdate} -- {event_name}",body)

    #@tasks.loop(minutes=60)
    async def Check4Application():
        # check inbox for any application emails
        # send a preliminary email asking basic questions, CCing Maria and James
        pass

    async def Check4Requests():                                                 # Checks for Appearance Request Forms
        # check inbox for any email with title "WOLFIE APPEARANCE REQUEST FORM"
        # turn all request emails into a dict containing all the required boxes to fill on sheet
        # post the request to the sheet
        # send out an email blast notifying wolfies that they can sign up
        # allow signup by replying to the email?
        # make sure email is after end date of the last weekly email that was sent
        nl = '\n'
        try:
            m = gmail.search_messages(service, "subject:WOLFIE APPEARANCE REQUEST FORM Submission", "in:unread in:inbox")
        except:
            m = []
        for msg in m:
            msg_dict = gmail.read_message(service, msg)
            body = msg_dict["body"].split("*NOTE: DO NOT REPLY * to this email as the address that it is sent from")[-1]
            print("+------------------------+")
            body = body.split(nl)
            requestform = []
            form_headers = ["Organization ", "Contact Name ","Address ","City ","State ","Zip ","Phone # ","Phone (Cell) ","Email ","Event Title ","Description ","Role of mascot at event ","Date of event","Time of requested appearance","Location","Driving Directions"]
            for x in body:
                try:
                    if x.split()[0] == ">":
                        x = x.replace("\r","").replace("> ","")
                        for y in form_headers:
                            if x.startswith(y):
                                requestform.append(x)
                                print(x)
                                break
                except: pass
            print(requestform)
            reqdate = requestform[12].replace("Date of event ","").split("/")
            print(reqdate)
            day = reqdate[1]
            month = reqdate[2]
            reqdate = reqdate[0] + "/" + reqdate[1]
            print(reqdate)
            cellorder = [
                await utils.wkday(reqdate),
                reqdate,
                requestform[9].replace("Event Title ",""),
                requestform[14].replace("Location ",""),
                requestform[13].replace("Time of requested appearance ",""),
                requestform[11].replace("Role of mascot at event ",""),
                1, "", "",
                requestform[1].replace("Contact Name ",""),
                requestform[8].replace("Email ",""),
                (requestform[6].replace("Phone #","") + nl + requestform[7].replace("Phone (Cell)","")).replace(" ",""),
                datetime.today().strftime(f"#{monthvar}m/%d"),
            ]
            print("--------------------------------------")
            val = (datetime.strptime(reqdate, "%m/%d") - timedelta(days=1)).strftime(f"#{monthvar}m/%d")
            while True:
                print(val)
                try:
                    eventb4cell = pygsheets.Worksheet.find(wolfie_schedule,val,cols=(2,2),matchEntireCell=True)[-1]
                    break
                except:
                    val = (datetime.strptime(val, "%m/%d") - timedelta(days=1)).strftime(f"#{monthvar}m/%d")
            eventdate = await utils.GetDate(reqdate)
            earliestdate = await utils.GetDate(weekapi.cell("A4").value)
            wolfie_schedule.insert_rows(eventb4cell.row,1)
            row = eventb4cell.row + 1
            a = 1
            for x in cellorder:
                wolfie_schedule.update_value((row,a),x)
                a += 1
            today = datetime.strptime(datetime.today().strftime("%b %d %Y"), "%b %d %Y")
            if eventdate > earliestdate or eventdate - today > timedelta(days=event_confirmation_days):
                if eventdate > earliestdate:
                    gmail.send(service, Management_Emails,f'Appearance Request Form Processed ({requestform[9].replace("Event Title ","")})',f'The request form sent by {msg_dict["from"]} was processed and added to the spreadsheet.')
                elif eventdate - today > timedelta(days=event_confirmation_days):
                    gmail.send(service, await utils.GetEmailList(),f'NEW WOLFIE SHIFT POSTED -- {reqdate} -- {requestform[9].replace("Event Title ","")}',f'A shift on {await utils.wkday(reqdate)} {reqdate} has been posted!{nl}{nl}Event Location: {requestform[14].replace("Location ","")}{nl}Event Time: {requestform[13].replace("Time of requested appearance ","")}{nl}{nl}Reply to this email to sign up!')
                    gmail.send(service, Management_Emails,f'Appearance Request Form Processed ({requestform[9].replace("Event Title ","")})',f'The request form sent by {msg_dict["from"]} was processed and added to the spreadsheet.')
            else:
                gmail.send(service, await utils.GetEmailList(),f'Appearance Request Form Processing Failed ({requestform[9].replace("Event Title ","")})',f'The request form sent by {msg_dict["from"]} has too quick of a turn around to be processed and added to the spreadsheet.')
        try: gmail.mark_as_read(service,"in:inbox subject:WOLFIE APPEARANCE REQUEST FORM Submission")
        except: pass

    async def EmailConfirmation():                                              # Checks for responses to sent emails
        nl = '\n'
        try:
            m = gmail.search_messages(service, "subject:CONFIRM WOLFIE SHIFT ", "in:unread in:inbox")
        except:
            m = []
        for msg in m:
            msg_dict = gmail.read_message(service, msg)
            title = msg_dict["subject"].split(" -- ")
            dates = pygsheets.Worksheet.find(wolfie_schedule,f'{title[1]}',cols = (1,2))
            for event in dates:
                print(event)
                row = event.row
                c1 = wolfie_schedule.cell((row,3))
                print(c1.value)
                if title[2].lower() in c1.value.lower():
                    print("Event Found!")
                    sender = msg_dict["from"].split("<")[-1].replace(">","")
                    print(sender)
                    from_cell = pygsheets.Worksheet.find(contact_info,sender,cols = (5,7),matchCase=False)
                    print(from_cell)
                    from_row = from_cell[0].row
                    from_name = contact_info.cell(f"A{from_row}").value
                    final_cell = pygsheets.Worksheet.find(wolfie_schedule,f'{from_name}',cols = (8,9),rows=(row,row),matchCase=False)
                    print(final_cell)
                    final_col = final_cell[0].col
                    if final_col == 8: wolfie_schedule.update_value(f"O{row}", "x")
                    elif final_col == 9: wolfie_schedule.update_value(f"P{row}", "x")
                    gmail.send(service,sender,f"Shift Confirmation Recieved!",f'Thank you for confirming your shift for "{title[2]}"{nl}{nl}Contact {Marketing_Director} or {Project_Maintainer} if you have a sudden availability change.')
        try:
            gmail.mark_as_read(service,"subject:CONFIRM WOLFIE SHIFT ")
            print("Confirmations cleared")
        except:
            pass
        # ------------------------------- #
        try:
            m = gmail.search_messages(service, "subject:WOLFIE EVENTS WEEK OF ", "in:unread in:inbox")
        except:
            m = []
        for msg in m:
            msg_dict = gmail.read_message(service, msg)
            title = msg_dict["subject"].split(" -- ")
            date_range = title[1].split(" - ")
            if date_range[0] == weekapi.cell("A2").value:
                print(title)
                print(date_range)
                print(msg_dict)
                email_body = msg_dict["body"].split("<wolfie.seawolf.bot@gmail.com> wrote:\r\n\r\n> The Following events need to be filled this week:")[0]
                email_body = list(email_body)
                print(email_body)
                for char in email_body:
                    try:
                        char = int(char)
                        event = week_events[char][3]
                        print(event)
                        print(char)
                        final_event = pygsheets.Worksheet.find(wolfie_schedule,event,cols=(3,3))
                        print(final_event)
                        sender_email = msg_dict["from"].split("<")[-1].replace(">","")
                        print(sender_email)
                        sender_cell = pygsheets.Worksheet.find(contact_info,sender_email,cols=(6,6))[0]
                        sendername = contact_info.cell(f'A{sender_cell.row}').value
                        sender = sendername.replace(" ","")
                        print(sender)
                        row = final_event[0].row
                        wolfie = wolfie_schedule.cell(f"H{row}").value
                        spotters = wolfie_schedule.cell(f"I{row}").value
                        wolfies = wolfie.replace(" ","").split(",")
                        spotters = spotters.replace(" ","").split(",")
                        staff = wolfies + spotters
                        print(staff)
                        wolfcount = int(wolfie_schedule.cell(f"G{row}").value)
                        if sender not in staff:
                            if len(wolfies) != wolfcount:
                                wolfie_schedule.update_value(f"H{row}",wolfie + ", " + sendername)
                                gmail.send(service,sender_email,"Wolfie Shift Request Approved",f'You have been confirmed as a WOLFIE for the event "{final_event[0].value}"{nl}{nl}If you have a change in availability you can remove yourself from this event on the spreadsheet.')
                            elif wolfie_schedule.cell(f"I{row}").value == "":
                                wolfie_schedule.update_value(f"I{row}", sendername)
                                gmail.send(service,sender_email,"Wolfie Shift Request Approved",f'You have been confirmed as a SPOTTER for the event "{final_event[0].value}"{nl}{nl}If you have a change in availability you can remove yourself from this event on the spreadsheet.')
                            else:
                                gmail.send(service,sender_email,"Wolfie Shift Request Denied",f'The event "{final_event[0].value}" appears to be filled! If you believe this is a mistake, you can check the spreadsheet.{nl}{nl}If this was a mistake, please reach out to {Project_Maintainer} so he can fix my code!')
                        else:
                            gmail.send(service, sender_email,"Wolfie Shift (Already Assigned)",f'At the time of your request, you were already assigned to the event "{final_event[0].value}". Looking forward to seeing you there!{nl}{nl}If this is an error, please reach out to {Project_Maintainer} so he can fix my code')
                        break
                    except:
                        raise
        try:
            gmail.mark_as_read(service, "in:unread in:inbox subject:WOLFIE EVENTS WEEK OF ")
            print("Week of emails cleared")
        except:
            pass
        # ------------------------------- #
        try:
            m = gmail.search_messages(service, "subject:NEW WOLFIE SHIFT POSTED ", "in:unread in:inbox")
        except:
            m = []
        for msg in m:
            msg_dict = gmail.read_message(service, msg)
            title = msg_dict["subject"].split(" -- ")
            dates = pygsheets.Worksheet.find(wolfie_schedule,f'{title[1]}',cols = (1,2))
            for event in dates:
                print(event)
                row = event.row
                c1 = wolfie_schedule.cell((row,3))
                print(c1.value)
                if title[2].lower() in c1.value.lower():
                    print("Event Found!")
                    sender = msg_dict["from"].split("<")[-1].replace(">","")
                    print(sender)
                    from_cell = pygsheets.Worksheet.find(contact_info,sender,cols = (5,7),matchCase=False)
                    print(from_cell)
                    from_row = from_cell[0].row
                    from_name = contact_info.cell(f"A{from_row}").value
                    print(from_name)
                    wolfie = wolfie_schedule.cell(f"H{row}").value
                    spotter = wolfie_schedule.cell(f"I{row}").value
                    wolfies = wolfie.replace(" ","").split(",")
                    spotters = spotter.replace(" ","").split(",")
                    wolfcount = int(wolfie_schedule.cell(f"G{row}").value)
                    staff = wolfies + spotters
                    print(staff)
                    if from_name.replace(" ","") not in staff:
                        if len(wolfies) != wolfcount:
                            if wolfies[0] == "": string = from_name
                            else: string = wolfie + ", " + from_name
                            wolfie_schedule.update_value(f"H{row}",string)
                            gmail.send(service,sender,"Wolfie Shift Request Approved",f'You have been confirmed as a WOLFIE for the event "{c1.value}"{nl}{nl}If you have a change in availability you can remove yourself from this event on the spreadsheet.')
                        elif wolfie_schedule.cell(f"I{row}").value == "":
                            wolfie_schedule.update_value(f"I{row}", from_name)
                            gmail.send(service,sender,"Wolfie Shift Request Approved",f'You have been confirmed as a SPOTTER for the event "{c1.value}"{nl}{nl}If you have a change in availability you can remove yourself from this event on the spreadsheet.')
                        else:
                            gmail.send(service,sender,"Wolfie Shift Request Denied",f'The event "{c1.value}" appears to be filled! If you believe this is a mistake, you can check the spreadsheet.{nl}{nl}If this was a mistake, please reach out to {Project_Maintainer} so he can fix my code!')
                    else:
                        gmail.send(service, sender,"Wolfie Shift (Already Assigned)",f'At the time of your request, you were already assigned to the event "{c1.value}". Looking forward to seeing you there!{nl}{nl}If this is an error, please reach out to {Project_Maintainer} so he can fix my code')
        try:
            gmail.mark_as_read(service, "in:unread in:inbox subjectNEW WOLFIE SHIFT POSTED ")
            print("New shift signups cleared")
        except:
            pass
#OLD ON READY
#x = 1
#    date = "1"
#    return
#    print("looping")
#    while date != "":
#        date = weekapi.cell(f"B{x}").value
#        time = weekapi.cell(f"C{x}").value
#        name = weekapi.cell(f"E{x}").value
#        loca = weekapi.cell(f"D{x}").value
#        week_events[x] = ["CHK-SHT",date, time, name, loca]
#        print(x)
#        x += 1
#    print(week_events)
#    await main()
#    WolfieAutomation.LookForNewEmails.start()
#    WolfieAutomation.WeeklyEventEmail.start()
#    WolfieAutomation.DailyCheck.start()
