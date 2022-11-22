from secret import adr, Account
import smtplib
import pandas as pd
import re


def auth(credentials: Account) -> (smtplib.SMTP, str):
    server = smtplib.SMTP(credentials.smtp_server)
    server.starttls()
    server.login(credentials.login, credentials.password)
    return server, credentials.email


def send_mail(from_email: str, to_email: list[str], subject: str, email_message: str,
              server: smtplib.SMTP) -> dict[str, tuple[int, bytes]]:
    payload = f"From: {from_email}\r\nTo: {', '.join(to_email)}\r\nSubject: {subject}\r\n\n{email_message}\n"
    return server.sendmail(from_email, to_email, payload)


def read_file(path: str) -> str:
    with open(path, 'r') as f:
        return f.read()


def main():
    message = read_file('data/message')
    df = pd.read_excel('data/run_results-20.xltx')
    authenticated_server, sender_email = auth(adr)
    counting_company_outreaches = {company: 0 for company in df['Company']}
    log = []
    for index, row in df.iterrows():
        if counting_company_outreaches[row['Company']] < 5:
            counting_company_outreaches[row['Company']] += 1
            person_emails = [row.Email1]  # , row.Email2, row.Email3]
            df.loc[index, 'Outreach'] = True
            for email in person_emails:
                print(f'Sending message to {email}')
                email = re.sub(r"@.*", lambda m: m.group().lower(), email)
                try:
                    send_mail(sender_email, [email], 'AI for code quality analytics',
                              message.format(first_name=row['Name'], company_name=row['Company']), authenticated_server)
                    df.loc[index, 'Sent'] = True
                except Exception:
                    log.append(f'Failed to send message to {email}')

    pd.DataFrame(counting_company_outreaches.items(), columns=['Company', 'Outreaches'])\
        .to_excel('data/companies.xlsx', index=False)
    df.to_excel('data/after_run.xlsx', index=False)
    with open('data/logs', 'w+') as f:
        f.write('\n'.join(log))
    authenticated_server.quit()


if __name__ == "__main__":
    main()
