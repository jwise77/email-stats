from imaplib import IMAP4_SSL
import os, sys, traceback, getpass
from datetime import datetime
from time import mktime
import email
import numpy as np
import matplotlib.pyplot as plt
from matplotlib import dates, ticker, patches
try:
    import cPickle as pickle
except:
    import pickle

imap = None
pickle_fn = 'headers.cpkl'

def parse_mailbox(data):
    flags, b, c = data.partition(' ')
    separator, b, name = c.partition(' ')
    return (flags, separator.replace('"', ''), name.replace('"', ''))

def fetch_headers():
    # Get server information
    host = raw_input("Server (e.g. imap.gmail.com): ")
    user = raw_input("Username: ")
    password = getpass.getpass("Password: ")
    folder = raw_input("Folder (recursive): ")
    try:
        # Create the IMAP Client
        imap = IMAP4_SSL(host)
        # Login to the IMAP server
        resp, data = imap.login(user, password)
        all_times = []
        if resp == 'OK' and bytes.decode(data[0]) == 'LOGIN completed.':
            totalMsgs = 0
            # Get a list of mailboxes
            resp, data = imap.list('"{0}"'.format(folder), '*')
            if resp == 'OK':
                for mbox in data:
                    flags, separator, name = parse_mailbox(bytes.decode(mbox))
                    # Select the mailbox (in read-only mode)
                    imap.select('"{0}"'.format(name), True)
                    # Get ALL message numbers
                    resp, msgnums = imap.search(None, 'ALL')
                    mycount = len(msgnums[0].split())
                    totalMsgs = totalMsgs + mycount
                    if mycount == 0: continue
                    uids = msgnums[0].replace(' ', ',')
                    resp, raw_headers = imap.fetch(uids, '(BODY.PEEK[HEADER.FIELDS (Date)])')
                    for header in raw_headers:
                        if header == ')': continue
                        em = email.message_from_string(header[1])
                        all_times.append(em['date'])
                    print('{:<30} : {: d}'.format(name, mycount))
            print('{:<30} : {: d}'.format('TOTAL', totalMsgs))
    except:
        print('Unexpected error : {0}'.format(sys.exc_info()[0]))
        traceback.print_exc()
    finally:
        if imap != None:
            imap.logout()
        imap = None
    return all_times

def load_headers():
    if os.path.exists(pickle_fn):
        fp = open(pickle_fn, 'rb')
        time_data = pickle.load(fp)
        fp.close()
    else:
        time_strings = fetch_headers()
        nmsg = len(time_strings)
        time_data = dict(unix_time = np.zeros(nmsg),
                         hour = np.zeros(nmsg),
                         day = np.zeros(nmsg),
                         date = np.zeros(nmsg),
                         month = np.zeros(nmsg),
                         doy = np.zeros(nmsg))
        for i,ts in enumerate(time_strings):
            time_tuple = email.utils.parsedate_tz(ts)
            unix_time = email.utils.mktime_tz(time_tuple)
            time = datetime.fromtimestamp(unix_time)
            time_data['unix_time'][i] = unix_time
            time_data['hour'][i] = time.hour
            time_data['day'][i] = time.weekday()
            time_data['date'][i] = time.day
            time_data['month'][i] = time.month
            time_data['doy'][i] = time.timetuple().tm_yday
        fp = open(pickle_fn, 'wb')
        pickle.dump(time_data, fp)
        fp.close()
    return time_data

def running_mean(x, N):
    cs = np.cumsum(np.insert(x,0,0))
    return (cs[N:] - cs[:-N]) / N

def process_data(data):
    """
    Various plots of time/date tendencies 
    """
    years = dates.YearLocator()
    months = dates.MonthLocator()
    yearsFmt = dates.DateFormatter('%Y')
    dateconv = np.vectorize(datetime.fromtimestamp)

    #
    # Complete history by day with weekly/monthly running averages
    #
    fig, ax = plt.subplots()
    # Floor to midnight (strip off time)
    dd = datetime.fromtimestamp(data['unix_time'].min()).date()
    start_time = mktime(dd.timetuple())  # Convert back to unix time
    # Ceiling to 23:59 of the last email
    dd = datetime.fromtimestamp(data['unix_time'].max()).date()
    end_time = mktime(dd.timetuple()) + 86359.0  # Convert back to unix time
    ndays = int(np.ceil((end_time - start_time) / 86400))
    # Bin by dates
    H, edges = np.histogram(data['unix_time'], range=(start_time, end_time),
                            bins=ndays)
    dcenter = 0.5*(edges[:-1] + edges[1:])
    plt.plot(dateconv(dcenter), H, lw=0.5, color='k', label='daily')
    if ndays > 7:
        plt.plot(dateconv(dcenter[6:]), running_mean(H,7),
                 lw=3, color='b', label='7-day')
    if ndays > 30:
        plt.plot(dateconv(dcenter[29:]), running_mean(H,30),
                 lw=3, color='r', label='30-day')
    plt.xlabel('Time')
    plt.ylabel('Emails / day')
    plt.legend(loc='best')
    plt.subplots_adjust(left=0.03, right=0.99, bottom=0.13, top=0.97)
    fig.set_size_inches(20,4)
    plt.savefig('history.png')

    #
    # Binned by hour
    #
    width = 0.8
    fig, ax = plt.subplots()
    H, edges = np.histogram(data['hour'], range=(-0.5,23.5), bins=24)
    plt.bar(np.arange(24)+0.5-width/2, H/float(ndays), width, color='grey', ec='k')
    plt.xlim(-0.1, 24.1)
    ax.xaxis.set_major_locator(ticker.MultipleLocator(1))
    plt.xlabel('Hour')
    plt.ylabel('Emails / day / hour')
    plt.subplots_adjust(left=0.1, right=0.95, bottom=0.1, top=0.95)
    plt.savefig('hourly.png')

    #
    # Binned by day of the week
    #
    fig, ax = plt.subplots()
    H, edges = np.histogram(data['day'], range=(-0.5,6.5), bins=7)
    # Move Sunday to the beginning of the week
    sunday = H[-1]
    H[1:] = H[:-1]
    H[0] = sunday
    plt.bar(np.arange(7)+0.5-width/2, H/(ndays/7.0), width, color='grey', ec='k')
    plt.xlim(-0.1, 7.1)
    plt.xticks(np.arange(7)+0.5, ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'])
    plt.ylabel('Emails / day')
    plt.subplots_adjust(left=0.08, right=0.95, bottom=0.08, top=0.95)
    plt.savefig('daily.png')

    #
    # Binned by date of the month
    #
    fig, ax = plt.subplots()
    H, edges = np.histogram(data['date'], range=(0.5,31.5), bins=31)
    plt.bar(np.arange(31)+0.5-width/2, H/(ndays/30.), width, color='grey', ec='k')
    plt.xlim(-0.1, 31.1)
    plt.xticks(np.arange(31)+0.5, ['%d' % (i+1) for i in range(31)])
    ax.tick_params(axis='x', which='both', labelsize=10)
    #ax.xaxis.set_major_locator(ticker.MultipleLocator(1))
    plt.xlabel('Date of Month')
    plt.ylabel('Emails / day')
    plt.subplots_adjust(left=0.1, right=0.95, bottom=0.1, top=0.95)
    plt.savefig('by-date.png')

    #
    # Binned by week of the year
    #
    width = 0.7
    fig, ax = plt.subplots()
    ax1 = ax.twiny()
    week = np.minimum(data['doy']/7.0, 52)
    H, edges = np.histogram(week, range=(0,52), bins=52)
    plt.bar(np.arange(52)+0.5-width/2, H/(ndays/52.), width, color='grey', ec='k')
    plt.xlim(-0.1, 52.1)
    # Bottom y-axis: week of the year
    # Top y-axis: Month
    days_in_month = [0,31,28,31,30,31,30,31,31,30,31,30,31]
    approx_week = np.cumsum(days_in_month) / 7.0
    month_center = 0.5*(approx_week[:-1] + approx_week[1:])
    months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug',
              'Sep', 'Oct', 'Nov', 'Dec']

    ax1.set_xticks(approx_week[:-1])
    ax1.set_xlim(-0.1, 52.1)
    ax1.set_xticklabels(months, ha='left')
    ax.xaxis.set_major_locator(ticker.MultipleLocator(5))
    ax.set_xlim(-0.1, 52.1)
    ax.set_ylabel('Emails / day')
    ax.set_xlabel('Week of the year')
    plt.subplots_adjust(left=0.08, right=0.97, bottom=0.1, top=0.92)
    plt.savefig('weekly.png')

    #
    # Punchcard graph
    # Inspired by https://www.mercurial-scm.org/wiki/PunchcardExtension
    #
    fig, ax = plt.subplots()
    H, xe, ye = np.histogram2d(data['hour'], data['day'], bins=(24,7),
                               range=((0,24),(-0.5,6.5)))
    # Move Sunday to the first day of the week
    sunday = H[:,-1]
    H[:,1:] = H[:,:-1]
    H[:,0] = sunday
    day_hour = np.mgrid[0:24, 0:7]
    norm = 200.0 / H.max()
    plt.scatter(day_hour[0].ravel(), day_hour[1].ravel(), s=norm*H.ravel(),
                color='k')
    # Mark typical working day
    rect = patches.Rectangle((8.5,0.5), 9, 5, color='c', alpha=0.3)
    ax.add_patch(rect)
    # Setup custom ticks
    ax.xaxis.set_major_locator(ticker.MultipleLocator(1))
    plt.yticks(range(7), ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'])
    plt.xlabel('Hour')
    plt.xlim(-0.5, 23.5)
    plt.ylim(6.5, -0.5)
    plt.subplots_adjust(left=0.08, right=0.97, bottom=0.1, top=0.97)
    plt.savefig('punchcard.png')
    

if __name__ == "__main__":
    data = load_headers()
    process_data(data)

# resp, raw_header = imap.fetch(num, '(BODY[header])')
# header = email.message_from_string(raw_header)
# time_parts = email.utils.parsedate(header['date'])
