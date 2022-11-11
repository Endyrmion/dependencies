import getpass
from collections import defaultdict
from exchangelib import Account, Configuration, Credentials, DELEGATE

def connect(server, email, username, password):
    creds = Credentials(username=username, password=password)
    config = Configuration(server=server, credentials=creds)
    return Account(primary_smtp_address=email, autodiscover=False, config = config, access_type=DELEGATE)

def print_tree(account):
    print(account.root.tree())

def get_recent_emails(account, folder_name, count):
    # Get the folder object
    folder = account.root / 'Top of Information Store' / folder_name
    # Get emails
    return folder.all().order_by('-datetime_received')[:count]

def count_senders(emails):
    counts = defaultdict(int)
    for email in emails:
        counts[email.sender.name] += 1
    return counts

def print_non_replies(emails, agents):
    dealt_with = dict()
    not_dealt_with = dict()
    not_dealt_with_list = list()
    for email in emails:
        subject = email.subject.lower().replace('re: ', '').replace('fw: ', '')

        if subject in dealt_with or subject in not_dealt_with:
            continue
        elif email.sender.name in agents:
            dealt_with[subject] = email
        else:
            not_dealt_with[subject] = email
            not_dealt_with_list += [email.subject]

    print('NOT DEALT WITH:')
    for subject in not_dealt_with_list:
        print(' * ', subject)

def main():
    server = 'outlook.office365.com'
    email = 'youness.ayouch@stonehagefleming.com'
    username = 'youness.ayouch@stonehagefleming.com'
    password = getpass.getpass()
    account = connect(server, email, username, password)
    emails = get_recent_emails(account, 'Inbox', 50)

    # agents = {
    #    'John':True,
    #    'Jane':True,
    #}
    # print_non_replies(emails, agents)

if __name__ == '__main__':
    main()