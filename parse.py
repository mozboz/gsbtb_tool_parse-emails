import mailbox
from email.header import Header, decode_header, make_header
import csv
#import cgi
#from StringIO import StringIO as IO
import re

subjectsToParse = ['Project Application', 'New Project', 'New Expat', 'Expat Contact']
csvQuoteBehaviour = csv.QUOTE_MINIMAL
csvColumns = ['ID', 'date', 'subject', 'labels', '']
#cgi.cgitb.enable()
regexp = re.compile('(.*)\nname: (.*)\nsurname: (.*)\nemail: (.*)\nphone: (.*)\nabout: ([\s\S]*)\nrelation: (.*)\nrelationname: (.*)')

def make_it_utf8(value):
    # seems to work with Python 3.3, cf. http://stackoverflow.com/questions/7331351/python-email-header-decoding-utf-8/21715870#21715870
    #return make_header(decode_header(value))
    # but the follow seems to do the job, too
    return ' '.join((item[0].decode(item[1] or 'utf-8').encode('utf-8') for item in decode_header(value)))

def message_body(message):#, boundary):
    # first item only, ignore multipart attachements
    content = message.get_payload(0).get_payload(0)
    content = str(content).replace('\r\n', '\n')
    result = regexp.search(content)
    if result:
        return [result.group(index).strip() for index in range(2, 9)]
    else:
        print "Could not parse content of"
        print message


def to_cvs(inboxFile):
    countTotal = 0
    countParsed = 0
    with open('mails.csv', 'wb') as csvfile:
        writer = csv.writer(csvfile, quoting=csvQuoteBehaviour)
        writer.writerow(csvColumns) # write a header line
        inbox = mailbox.mbox(inboxFile)
        print "Messages in file: {0}".format(len(inbox))

        for message in inbox:
            countTotal += 1
            if make_it_utf8(message['subject']) in subjectsToParse:
                countParsed += 1
                data = []

                # Extracting relevant information from this email

                # collect header data: message ID, subject, gmail labels
                data.append(message['message-Id'])
                data.append(message['date'])
                data.append(message['subject'])
                data.append(message['x-Gmail-Labels'])

                # parse content
                data.extend(message_body(message))

                data = [make_it_utf8(cell) for cell in data]
                writer.writerow(data)

    print 'stats: total {0}, parsed {1}'.format(countTotal, countParsed)
