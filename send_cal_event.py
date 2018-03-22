# -*- coding: utf-8 -*-
"""
Created on Fri Sep 26 08:53:48 2014

@author: lkonis
"""
import sys
import win32com.client
import datetime

OUTLOOK_APPOINTMENT_ITEM  = 1
OUTLOOK_MEETING           = 1
OUTLOOK_ORGANIZER         = 1
OUTLOOK_OPTIONAL_ATTENDEE = 2

ONE_HOUR       = 60
THIRTY_MINUTES = 30
TWO_DAYS       = ONE_HOUR*24*2
LOCATION       = 'Meating room'
SUBJECT        = 'Algo breakfast'
BODY           = 'It is your turn to bring the algo-team breakfast :-)'
SENDERNAME     = '\n\n<This is an automatic mail sent to you by Algo-Breakfast Robot>'#'Lior'
#OUTLOOK_FORMAT = '%m/%d/%Y %H:%M'
OUTLOOK_FORMAT = '%d-%m-%Y %H:%M:%S'

MONDAY      = 0
TUESDAY     = 1
WEDNESDAY   = 2
THURSDAY    = 3
FRIDAY      = 4


outlook_date   = lambda dt: dt.strftime(OUTLOOK_FORMAT)



class OutlookClient(object):
    def __init__(self):
        self.outlook = win32com.client.Dispatch('Outlook.Application')

    def next_weekday(self, d, weekday):# 0 = Monday, 1=Tuesday, 2=Wednesday..
        days_ahead = weekday - d.weekday()
        if days_ahead < 0: # Target day already happened this week
            days_ahead += 7
        return d + datetime.timedelta(days_ahead)

    def send_meeting_request(self, recipients=None, sender=None, send_date=None, name=None,desired_weekday=WEDNESDAY):
        args = sys.argv[1:]
        if (not args) & (send_date == None):
            print '\nsend_cal_event.py - a python script for sending breakfast invitation'
            print '\nusage: send_cal_event.py recipient-mail[,send_date (%d/%m/%y)[,organizer mail[,[recipient name]]]\n'
            raw_input ('press any key...\n')
            sys.exit(1)
        else:
            if recipients==None:
                if len(args)<2:
                    recipients = [args[0]]
                    sender='lkonis@gnresound.com'
                    send_date= datetime.date.today()
                    name=recipients[0]
                else:
                    recipients = [args[0]]
                    send_date= datetime.datetime.strptime(args[1], "%d/%m/%Y")
                    if len(args)>2:
                        sender=args[2]
                    if len(args)>3:
                        name=args[3]

        mtg = self.outlook.CreateItem(OUTLOOK_APPOINTMENT_ITEM)
        mtg.MeetingStatus = OUTLOOK_MEETING
        mtg.Location = LOCATION
        time = datetime.time(9,0,0)
        if send_date:
            next_weekday = self.next_weekday(send_date, desired_weekday)
            when = datetime.datetime.combine(next_weekday,time)
        else:
            today = datetime.date.today()
            next_weekday = self.next_weekday(today, desired_weekday)
            when = datetime.datetime.combine(next_weekday,time)

        if sender:
            # Want to set the sender to an address specified in the call
            # This is the portion of the code that does not work
            organizer      = mtg.Recipients.Add(sender)
            organizer.Type = OUTLOOK_ORGANIZER
        for recipient in recipients:
            #olRecipient = self.outlook.CreateRecipient(recipient)
            invitee      = mtg.Recipients.Add(recipient)
            invitee.Type = OUTLOOK_OPTIONAL_ATTENDEE
        if name:
            body = "Hi " + name + ",\n" + BODY
        else:
            name = ""
            body = "Hi,\n" + BODY

        body += "\nBR\n "+ SENDERNAME
            
        mtg.Subject                    = SUBJECT
        mtg.Start                      = outlook_date(when)
        mtg.Duration                   = ONE_HOUR
        mtg.ReminderSet                = True
        mtg.ReminderMinutesBeforeStart = TWO_DAYS
        mtg.ResponseRequested          = True
        mtg.Body                       = body
        # wait for confirmation
        print "Sender: "+ name +"\nsends breakfast invitation at: "
        print when
        print "To: "+recipients[0]+"\n"
        answer = raw_input("Are you sure you want to send? [Y]")
        if "n" in answer:
            return
        print answer

        # now send
        mtg.Send()

if __name__ == "__main__":
    
    ol = OutlookClient()
    meeting_date =  datetime.date.today()
    meeting_time = datetime.datetime.now() + datetime.timedelta(hours=3)

    required_recipients = ['lior.konis@gmail.com']
    test_sender     = 'lkonis@gnresound.com'

    ol.send_meeting_request(required_recipients, test_sender, None, None)
