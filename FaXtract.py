"""
FaXtract by Eric W. Wallace/Atlantic Fund Administation
v0.3 EWW 2009-03-13 - added counter file and reporting
v0.2 EWW 2009-03-12 - added config file for multiple mailboxes
v0.1 EWW 2009-03-11 - proof-of-concept

Description: Connects to email mailboxes via IMAP4 and extracts attached fax PDF files.

Usage:
- Configure FaXtract.ini with the global options and mailboxes you want to check.
- Use command-line option '/report' to send an email report on the received faxes.
"""

__author__ = "Eric W. Wallace <e@ewall.org>"
__version__ = "$Revision: 0.3 $"
__date__ = "$Date: 2009/03/13 $"
__copyright__ = "Copyright (c)2009 Eric W. Wallace & Atlantic Fund Administration"
__license__ = "Creative Commons Attribution-ShareAlike (http://creativecommons.org/licenses/by-sa/3.0/)"

from datetime import datetime
import ConfigParser
import imaplib, email.parser, base64
import os.path, sys
import pickle

### Constants
VERBOSE = True
configFileName = 'FaXtract.ini'
counterFileName = 'FaXtract.cnt'
MYDATE = datetime.now().strftime('%Y%m%d%H%M%S')
REPORT = False #default value is overridden only if True

### Objects & Functions

def getFaxes(hostname, useSSL, username, password, prefix):
    """ process this mailbox and export the faxes """
    c = imap_connection(hostname, useSSL, username, password)
    processed = 0
    try:
        #open Inbox
        typ, data = c.select('INBOX', readonly=True)
        if VERBOSE: print "Inbox Status: " + typ, data
        if int(data[0]) == 0:
            if VERBOSE: print "No new email messages found."
            return 0

        #query for fax messages
        typ, msg_ids = c.search(None, SEARCHSTRING)
        if msg_ids[0] == '':
            if VERBOSE: print "No fax messages found."
            return 0
        msg_ids = msg_ids[0].split(" ") #make it a usable list
        if VERBOSE:
            print "Found", len(msg_ids), "fax messages."
            #print "Msg_ids:", msg_ids

        #loop thru messages
        for msg_id in msg_ids:
            typ, msg_data = c.fetch(msg_id, '(RFC822)')
            
            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_string(response_part[1])

                    if VERBOSE:
                        print "\nTIME    :", datetime.now()
                        for header in [ 'to', 'from', 'subject' ]:
                            print '%-8s: %s' % (header.upper(), msg[header])

                    if msg.is_multipart():
                        for m in msg.get_payload():
                            if (m.get_filename() != None) and (m.get_filename()[-4:].upper() == '.PDF'):
                                filename = prefix + "fax" + str(MYDATE) + "_" + str(msg_id) + ".pdf"
                                if VERBOSE:
                                    print "ATTACHED:", m.get_filename()
                                    print "FILENAME:", filename
                                file = open(os.path.join(ARCHIVESPATH, filename), 'wb')
                                file.write(base64.decodestring(m.get_payload()))
                                file.close()
                                file2 = open(os.path.join(EMSPATH, filename), 'wb')
                                file2.write(base64.decodestring(m.get_payload()))
                                file2.close()
                                # move message to "Processed" folder
                                typ, response = c.create(DONEFOLDER)
                                c.copy(msg_id, DONEFOLDER) #copy
                                typ, response = c.store(msg_id, '+FLAGS', r'(\Deleted)') #then delete
                                processed += 1
                            #else:
                            #    if VERBOSE: print "ATTACHED: Non-PDF attachments are ignored."
                    else:
                        if VERBOSE: print "Found email with no attachments; skipping it."
                                
    finally:
        try:
            c.close()
        except:
            pass
        c.logout()
    return processed

def sendReport():
    """ email a report of the number of faxes processed """
    # Read in the counts
    if os.path.exists(counterFileName):
        counterFile = open(counterFileName, 'rb')
        counts = pickle.load(counterFile)
        counterFile.close()
    else:
        #this shouldn't ever happen
        print "\nERROR: counterFile not found, cannot report!\n"
        sys.exit(4)
    if VERBOSE:
        print "\nCompiling and sending report..."
        print "counterFile exists, shows:",counts
    # Compose and send email
    fileDate = datetime.fromtimestamp(os.path.getctime(counterFileName)).strftime('%Y-%m-%d %H:%M:%S')
    strMessage = "FaXtract has processed the following faxes since " + fileDate + ":\n"
    for j,k in counts.iteritems():
        strMessage = strMessage + "- from " + str(j) + " processed " + str(k) +" faxes\n"
    sendmail(HOSTNAME, "FaXtract@afa.local", REPORTEMAIL, "Report from FaXtract", strMessage)
    # Delete the counter file
    os.remove(counterFileName)

def sendmail(hostname=None, sender='', to='', subject='', text=''):
    """ simple sendmail relay using built-in modules """
    import smtplib
    import email.Message, email.iterators, email.generator

    message = email.Message.Message()
    message["To"] = to
    message["From"] = sender
    message["Subject"] = subject
    message.set_payload(text)
    mailServer = smtplib.SMTP(hostname)
    smtpresult = mailServer.sendmail(sender, to, message.as_string())
    mailServer.quit()
    if smtpresult:
        errstr = ""
        for recip in smtpresult.keys():
            errstr = "Could not delivery mail to: %s\nServer said: %s\n%s\n%s" % (recip, smtpresult[recip][0], smtpresult[recip][1], errstr)
        raise smtplib.SMTPException, errstr
    else:
        if VERBOSE: print "Email sent."

def imap_connection(hostname, useSSL, username, password):
    """ open a connection to the IMAP4 server """
    if VERBOSE: print 'Connecting to IMAP server on', hostname + "..."

    if useSSL:
        connection = imaplib.IMAP4_SSL(hostname)
    else:
        connection = imaplib.IMAP4(hostname)
        
    if VERBOSE: print 'Logging in as', username
    connection.login(username, password)
    return connection

### Main
if __name__ == '__main__':

    # Redirect output
    sys.stdout = open('FaXtract.log', 'wa')
    sys.stderr = open('FaXtract_errors.log', 'wa')

    # Check if we need to send report
    if (len(sys.argv)>1) and (sys.argv[1]=="/report"):
        REPORT = True

    # Read global options from config file
    config = ConfigParser.RawConfigParser()
    config.read(configFileName)
    HOSTNAME = config.get('options', 'mailserver')
    USESSL = config.getboolean('options', 'usessl')
    SEARCHSTRING = config.get('options', 'searchstring')
    ARCHIVESPATH = config.get('options', 'archivespath')
    EMSPATH = config.get('options', 'emspath')
    DONEFOLDER = config.get('options', 'donefolder')
    REPORTEMAIL = config.get('options', 'reportemail')

    # Double-check globals
    if not (os.path.exists(EMSPATH)):
        print "\nERROR: emspath not found!\n"
        sys.exit(2)
    if not (os.path.exists(ARCHIVESPATH)):
        print "\nERROR: archivespath not found!\n"
        sys.exit(3)

    # Read mailboxes from config file and process them
    mailboxes=[]
    i = 1
    while True:
        if config.has_section(str(i)):
            mailboxes.append ( dict( [('username', config.get(str(i),'username') ),
                ('password', config.get(str(i),'password') ),
                ('prefix', config.get(str(i),'prefix') )] ) )
            # Process this mailbox now
            results = getFaxes(HOSTNAME, USESSL, mailboxes[i-1]['username'], mailboxes[i-1]['password'], mailboxes[i-1]['prefix'])
            if VERBOSE: print "\nProcessed",results,mailboxes[i-1]['prefix'],"fax messages."
            
            # Record the results
            if os.path.exists(counterFileName):
                #if we have a counter file, load the data structure
                counterFile = open(counterFileName, 'rb')
                counts = pickle.load(counterFile)
                counterFile.close()
                if VERBOSE: print "counterFile exists, shows:",counts
                if mailboxes[i-1]['prefix'] in counts:
                    counts[mailboxes[i-1]['prefix']] += results
                else:
                    counts[mailboxes[i-1]['prefix']] = results
            else:
                #if there is no counter file, create the data structure                
                counts = { mailboxes[i-1]['prefix'] : results }
            #now we save the new counts
            if VERBOSE: print "Saving new counterFile with:",counts
            counterFile = open(counterFileName, 'wb')
            pickle.dump(counts,counterFile)
            counterFile.close()

            #inc and loop
            i+=1
        else:
            break

    if REPORT: sendReport()

