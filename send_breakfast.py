# -*- coding: utf-8 -*-
"""
Created on Mon Sep 08 09:00:56 2014

@author: lkonis
"""
import pickle
import pandas as pd
import datetime as dt
from send_cal_event import *
import os
csv_fname = 'breakfast_list.csv'
algo_members = ['Johnny', 'Michael', 'Kim', 'Torben', 'Ditlev', 'Casper', 'Niels', 'Christian', 'No one']
algo_emails = ['jfredsgaard', 'mchristoffersen', 'ksberthelsen', 'tarnthansen', 'draboel', 'ctroelsen','noddershede','clykkegaard']
breakfast_day = WEDNESDAY



pickle_file = 'param.pkl'
algo_name_email =     dict(zip(algo_members, algo_emails))
PENDING_NONE = -1
PENDING_ACCEPTED = 0
DONE = 1
weekdays = {'MONDAY':MONDAY, 'TUESDAY':TUESDAY, 'WEDNESDAY':WEDNESDAY, 'THURSDAY':THURSDAY, 'FRIDAY':FRIDAY}
str_day_of_week = weekdays.keys()[weekdays.values().index(breakfast_day)]
declined_members = ['']

# initial tables, should be executed only once
def generate_member_table():


    # create DataFrame
    r = pd.read_csv(csv_fname)
    p = pd.DataFrame(data=r)
    d = pd.to_datetime(p.date,exact=True, format="%d-%m-%Y")
    p.date = d
    p['status'] = PENDING_NONE
    p.loc[p.date > dt.date.today(), ('status')] = PENDING_ACCEPTED
    p.loc[p.date <= dt.date.today(), ('status')] = DONE
    p.to_csv("breakfast_list.csv", index=None)
    #print p.tail(10)



def load_list():
    global breakfast_day, declined_members, str_day_of_week
    p = pd.read_csv(csv_fname)

    try:
        with open(pickle_file, 'r') as fi:  # Python 3: open(..., 'wb')
            [breakfast_day_str, declined_members] = pickle.load(fi)
            breakfast_day = int(breakfast_day_str)
            str_day_of_week = weekdays.keys()[weekdays.values().index(breakfast_day)]
    except Exception as e:
        print 'load list exception:', e
        if os.path.isfile(pickle_file):
            os.remove(pickle_file)



    # check all names are legal
    #names = p.Name
    #if len(set(names) ^ set(algo_members))>0:
    #    print "Some names from list not exist in algo-members"
    #    print set(names) ^ set(algo_members)
    #    #sys.exit(0)
    # TODO - check all names are legal (are from member list)
    d = pd.to_datetime(p.date,exact=True, format="%Y-%m-%d")
    p.date = d
    # update events that are in the past
    p.loc[(p.date < dt.datetime.today()) & (p.status != DONE), 'status'] = DONE
    return p


def save_list(p):
    p.to_csv(csv_fname, index=False)
    with open(pickle_file, 'w') as fo:  # Python 3: open(..., 'wb')
        pickle.dump([str(breakfast_day), declined_members], fo)

# query who from algo-members brought latest back in history (and need to be next)
def who_to_invite_next(p):
    global declined_members
    p_last = pd.DataFrame()
    for member in algo_members[:-1]:
        if member in declined_members:
            continue
        if p[(p.Name == member) & (p.status <= PENDING_ACCEPTED)].empty: # if no pending (accepted)
            tmp_serie = p[(p.Name == member) & (p.status == DONE)].sort().tail(1)
            p_last = p_last.append(tmp_serie, ignore_index=True)
    d = p_last.sort(columns='date').head(1)
    return d
def next_weekday(d, weekday):# 0 = Monday, 1=Tuesday, 2=Wednesday..
    days_ahead = weekday - d.dayofweek
    if days_ahead < 0: # Target day already happened this week
        days_ahead += 7
    return d + dt.timedelta(days_ahead)

def expand_list(p, empty=0):
    m = who_to_invite_next(p)
    last_date = p.sort(columns='date').tail(1).date
    next_date = last_date + dt.timedelta(days=1) # the booking script will look for the next breakfast-day following this last-day + 1
    next_date = next_date.iloc[-1]
    next_date = next_weekday(next_date, breakfast_day)  # that is, wednesday

    cur_date = m.date
    m['date'] = next_date
    if empty==1:
        m.status = PENDING_ACCEPTED
        m.Name=algo_members[-1]
        print "no one is bringing breakfast at", m.squeeze().date.strftime("%d-%m-%Y")
    else:
        m.status = PENDING_NONE
        print m.Name.values[0],"is next to bring breakfast at", m.squeeze().date.strftime("%d-%m-%Y")

    p = p.append(m)
    p = p.reset_index(drop=True)
    return p

def find_next_pending(p):
    for index, row in p.T.iteritems():
        if row.status==PENDING_NONE:
            return index
    return False

if __name__ == '__main__':
    args = sys.argv[1:]
    s = '1'
    while s!="":
        os.system('cls')
        if s=='':
            sys.exit(0)
        p = load_list()
        if s == '0':
            print 'Full history:\n'
            print p.ix[p['status']==DONE,'Name':'date']
            #sys.exit(0)
        elif s == '1':
            print 'Assigned breakfasts (accepted):\n\n', p.ix[(p.status == PENDING_ACCEPTED),'Name':'date']
            non_pending = p.ix[(p.status == PENDING_NONE), 'Name':'date']
            if len(non_pending)>0:
                print '\nNext candidate - invitation not sent yet:\n',non_pending
            #sys.exit(0)
        elif s in ['2','2a','2b']:
            any_pending = p[p.status == PENDING_NONE]
            if len(any_pending)>0:
                for index, row in p[p.status == PENDING_NONE].T.iteritems():
                    who = p.ix[index]
                    if s=='2a':
                        print who.Name, "is invited to bring breakfast at", who.date
                    else:
                        print "Wait!", who.Name, "is invited to bring breakfast at", who.date
                    ans = raw_input("Did he accept? (a-accept, d-declined, <CR> to back to menu: ")
                    if ans.lower() == 'a':
                        p.ix[index,'status']= PENDING_ACCEPTED
                        declined_members = ['']
                        save_list(p)
                    elif ans.lower() == 'd':
                        p.drop(index, inplace=True)
                        declined_members.append(who.Name)
                        save_list(p)
                    else:
                        break

            else:
                p = expand_list(p, s=='2b')
                save_list(p)
                print 'Assigned breakfasts (accepted):\n', p.ix[(p.status == PENDING_ACCEPTED),'Name':'date']
                non_pending = p.ix[(p.status == PENDING_NONE), 'Name':'date']
                if len(non_pending)>0:
                    print '\nNext candidate - not sent yet:\n',non_pending
            #sys.exit(0)
        elif s == '3':
            for index, row in reversed(list(p[p.status!=DONE].T.iteritems())):
                print p.ix[index,'Name'], "is assigned to bring breakfast at", p.ix[index,'date'], "\n"
                ans=raw_input("Delete: (y/n)?  [n]") or "n"
                if ans.lower() == 'y':
                    p.drop(index, inplace=True)
                    print "Don''t forget to cancel booking from outlook\n"
                else:
                    break
            #p=p[p.status!=PENDING_NONE]
            save_list(p)
            print 'Assigned breakfasts (accepted):\n', p.ix[(p.status == PENDING_ACCEPTED),'Name':'date']
            non_pending = p.ix[(p.status == PENDING_NONE), 'Name':'date']
            if len(non_pending)>0:
                print '\nNext candidate - not sent yet:\n',non_pending
        elif s == '4':
            m = find_next_pending(p)
            if m:
                print "next candidate to invite is: ", p.ix[m,'Name']
                ol = OutlookClient()

                recipient_name = p.iloc[m].Name
                recipient_mail = [algo_name_email[recipient_name] + '@gnresound.com']
                meeting_date = p.iloc[m].date
                meeting_time = dt.timedelta(hours=9)
                test_sender = 'lkonis@gnresound.com'
                r = ol.send_meeting_request(recipient_mail, test_sender, meeting_date, recipient_name, breakfast_day)

            else:
                print("You need the expand the list with more candidate before you try to invite someone")
        elif s == '5':
            print weekdays
            chs_d = raw_input('choose day:')
            try:
                int_day_ch = int(chs_d)
                assert (int_day_ch >= 0 and int_day_ch <= 4)
                breakfast_day = int_day_ch
                str_day_of_week = weekdays.keys()[weekdays.values().index(breakfast_day)]
                save_list(p)
            except Exception as e:
                print 'not a week day'
        else:
            sys.exit(0)
        print '======================\n'
        #print os.path.basename(os.path.basename(__file__) ), '\nUsage:\n'
        print '\nMenu:\n===='
        print '[0]: show all past breakfasts'
        print '[1]: show future assigned breakfasts'
        print '[2]: Add new entry to list. Week-day for next meeting: ' , str_day_of_week
        print '[2a]: Decline next candidate'
        print '[2b]: Add new EMPTY entry to list (e.g. if holiday)'
        print '[3]: Delete invitations'
        print '[4]: send invitation to next candidate'
        print '[5]: change day of week'
        s = raw_input('\n')
    sys.exit(0)


