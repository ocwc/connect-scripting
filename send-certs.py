import xlrd
import click
import emails
import env


class CertificateMailer(object):
    data = {}
    workbook = None

    def __init__(self, filename):
        self.workbook = xlrd.open_workbook(filename)
        self.parse_data()
        self.send_emails()

    def parse_data(self):
        sheet = self.workbook.sheet_by_name("presentations")
        data = {}
        for row_idx in range(1, sheet.nrows):
            data[row_idx] = [
                sheet.cell(row_idx, 1).value,
                sheet.cell(row_idx, 4).value,
            ]
        self.data = data

    def send_emails(self):
        for key, (name, email) in self.data.items():
            subject = "OE Global Conference 2020 - Certificate of Participation"
            body = "Dear {},<br /><br />Please find attached your Certificate of Participation.<br /><br />~~ OE Global Team".format(
                name
            )
            msg = emails.html(
                subject=subject,
                html=body,
                mail_from=("OE Global", "conference@oeglobal.org"),
            )
            msg.attach(
                filename="oeglobal-certificate.pdf",
                data=open("certs/certificate-{}.pdf".format(key), "rb"),
            )

            print(subject, body, key)
            try:
                print(key, name, email)
                response = msg.send(to=email, smtp=env.SMTP_SETTINGS)
                print(response.status_code)
            except:
                print("exception", key, name, email)
                print(response.status_code)
                continue


@click.command()
@click.option("--filename", help="Custom Certificate mailer from xlsx")
def cli(filename):
    CertificateMailer(filename)


if __name__ == "__main__":
    cli()
