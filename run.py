import argparse
import parse

argParser = argparse.ArgumentParser(description='reads given mbox file, filters mails generated form on gsbtb.org, and stores the results as CSV.')
argParser.add_argument('--add', '--append', default=False, help='adds data to existing CSV files. Otherwise, CSV files will be deleted before running the parsing process.')
argParser.add_argument('mbox_files', nargs='+', help='location of mbox file(s)')
args = argParser.parse_args()

if not args.add:
    # delete existing CSV files before starting
    import glob
    import os
    for file in glob.glob("*.csv"):
        os.remove(file)

inboxFiles = args.mbox_files
print 'Parsing mbox files {0} ...'.format(', '.join(inboxFiles))
parse.to_cvs(inboxFiles)

# test message body parser:
#import mailbox
#inbox = mailbox.mbox(inboxFile)
#message = inbox[0]
#print parse.message_body(message)
