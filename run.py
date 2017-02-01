import parse

inboxFile = './Takeout/Mail/Incoming Monthly-2017-January.mbox'

parse.to_cvs(inboxFile)

#import mailbox
#inbox = mailbox.mbox(inboxFile)
#message = inbox[0]
#print parse.message_body(message)
