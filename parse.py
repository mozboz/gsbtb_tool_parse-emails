import mailbox
from email.header import Header, decode_header, make_header
import csv
#import cgi
#from StringIO import StringIO as IO
import re

csvQuoteBehaviour = csv.QUOTE_ALL # alternative: csv.QUOTE_MINIMAL

configuration = {
    'Project Application': {
        'regexp': [re.compile('(.*)\nname: (.*)\nsurname: (.*)\nemail: (.*)\nphone: (.*)\nabout: ([\s\S]*)\nrelation: (.*)\nrelationname: (.*)')],
        'first_result': 2,
        'last_result': 8,
        'file_name': 'project-applications.csv',
        'header': ['ID', 'date', 'subject', 'labels', 'name', 'surname', 'email', 'phone', 'about', 'relation', 'relationname']
    },
    'New Project': {
        'regexp': [re.compile('(.*)\nname: (.*)\naddress: (.*)\nphone: (.*)\ncontact: (.*)\nemail: (.*)\nwebsite: (.*)\nabout: ([\s\S]*)\ntask-description: ([\s\S]*)\nwhen: (.*)\nrequirements: (.*)\ngerman-skills: (.*)\nsex: (.*)')],
        'first_result': 2,
        'last_result': 13,
        'file_name': 'new-project.csv',
        'header': ['ID', 'date', 'subject', 'labels', 'name', 'address', 'phone', 'contact', 'email', 'website', 'about', 'task-description', 'when', 'reuqirements', 'german-skills', 'sex']
    },
    'New Expat': {
        'regexp': [
            re.compile('(.*)\nname: (.*)\nsurname: (.*)\nemail: (.*)\nphone: (.*)\nwebsite: (.*)\ngerman-skills: (.*)[\nsex: ]*(.*)\nabout: ([\s\S]*)\nidea: ([\s\S]*)\n\*')],
        'first_result': 2,
        'last_result': 10,
        'file_name': 'new-expat.csv',
        'header': ['ID', 'date', 'subject', 'labels', 'name', 'surname', 'email', 'phone', 'website', 'sex (optional)', 'about', 'idea']
    },
    'Expat Contact': {
        'regexp': [re.compile('(.*)\nname: (.*)\nemail: (.*)\nphone: (.*)\nmessage: ([\s\S]*)\nrelation: (.*)\nrelationname: (.*)')],
        'first_result': 2,
        'last_result': 7,
        'file_name': 'expat-contact.csv',
        'header': ['ID', 'date', 'subject', 'labels', 'name', 'email', 'phone', 'message', 'relation', 'relationname']
    }
}
writers = {}

def make_it_utf8(value):
    # seems to work with Python 3.3, cf. http://stackoverflow.com/questions/7331351/python-email-header-decoding-utf-8/21715870#21715870
    #return make_header(decode_header(value))
    # but the follow seems to do the job, too
    return ' '.join((item[0].decode(item[1] or 'utf-8').encode('utf-8') for item in decode_header(value)))

def message_body(message, config):
    # first item only, ignore multipart attachements
    content = message.get_payload(0).get_payload(0)
    content = str(content).replace('\r\n', '\n')
    for regexp in config['regexp']:
        result = regexp.search(content)
        if result:
            return [result.group(index).strip() for index in range(config['first_result'], config['last_result']+1)]
    print "Could not parse content. Config/message"
    print config
    print message

def parse(message, config):
    data = []

    # Extracting relevant information from this email

    # collect header data: message ID, subject, gmail labels
    data.append(message['message-Id'])
    data.append(message['date'])
    data.append(message['subject'])
    data.append(message['x-Gmail-Labels'])

    # parse content
    data.extend(message_body(message, config))

    return [make_it_utf8(cell) for cell in data]

def to_cvs(inboxFile):
    count_total = 0
    count_parsed = 0
    inbox = mailbox.mbox(inboxFile)
    print "Messages in file: {0}".format(len(inbox))

    for message in inbox:
        count_total += 1
        if count_total % 100 == 0:
            print count_total
        message_type = make_it_utf8(message['subject'])
        if message_type in configuration:
            config = configuration[message_type]
            with open(config['file_name'], 'a') as csvfile:
                writer = csv.writer(csvfile, quoting=csvQuoteBehaviour)
                if not message_type in writers:
                    writer = writers[message_type] = writer
                    writer.writerow(config['header']) # write the header as configured
                data = parse(message, config)
                count_parsed += 1
                writer.writerow(data)

    print 'stats: total {0}, parsed {1}'.format(count_total, count_parsed)
