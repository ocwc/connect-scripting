import click
import requests
import xlrd
import arrow
from .env import API_KEY, API_USER, API_HOST

CATEGORIES = {
    "Applications of Open Education Practices/Open Pedagogy/Open Education Research": 26,
    "Connecting Open Education to Primary and Secondary (K-12) Education": 27,
    "Global Collaboration, Strategies, & Policies in Open Education": 28,
    "Innovation through MOOCs practices": 29,
    "Technologies for Open Education": 30,
}


class DiscourseImporter(object):
    def __init__(self, filename):
        self.workbook = xlrd.open_workbook(filename)
        self.create_topics()

    # def get_tags(self):
    #     current_tags = requests.get("{}/tags.json".format(API_HOST)).json().get("tags")
    #
    #     conf_tags = [
    #         "Applications of Open Education Practices/Open Pedagogy/Open Education Research",
    #         "Connecting Open Education to Primary and Secondary (K-12) Education",
    #         "Global Collaboration, Strategies, & Policies in Open Education",
    #         "Innovation through MOOCs practices",
    #         "Technologies for Open Education",
    #     ]
    #
    #     # for tag in current_tags:
    #     #     if tag['text'] in conf_tags:
    #     #         pass

    def _new_post(self, **kwargs):
        pass

    def create_topics(self):

        for sheetname, tz in [
            ("Taiwan-16-18-20", "Asia/Taipei"),
            # ("Netherlands-16-18-20", "Europe/Berlin"),
            # ("Canada-16-18-20", "America/Toronto"),
        ]:
            sheet = self.workbook.sheet_by_name(sheetname)
            for row_idx in range(3, sheet.nrows):
                session_format = sheet.cell(row_idx, 0).value
                timezone = sheet.cell(row_idx, 1).value
                easychair = int(sheet.cell(row_idx, 2).value)
                authors = sheet.cell(row_idx, 3).value
                title = sheet.cell(row_idx, 4).value
                sync = sheet.cell(row_idx, 5).value
                sector = sheet.cell(row_idx, 6).value
                unesco = sheet.cell(row_idx, 7).value
                topic = sheet.cell(row_idx, 8).value
                duration = sheet.cell(row_idx, 9).value
                duration = xlrd.xldate_as_tuple(duration, self.workbook.datemode)
                date = sheet.cell(row_idx, 10).value
                date = xlrd.xldate_as_tuple(date, self.workbook.datemode)
                start = sheet.cell(row_idx, 11).value
                start = xlrd.xldate_as_tuple(start, self.workbook.datemode)
                end = sheet.cell(row_idx, 12).value
                end = xlrd.xldate_as_tuple(end, self.workbook.datemode)
                try:
                    description = sheet.cell(row_idx, 13).value
                except IndexError:
                    description = None

                start_utc = (
                    arrow.get(*date, tzinfo=tz)
                    .replace(hour=start[3], minute=start[4], second=start[5])
                    .to("utc")
                )
                end_utc = (
                    arrow.get(*date, tzinfo=tz)
                    .replace(hour=end[3], minute=end[4], second=end[5])
                    .to("utc")
                )

                self._new_post(
                    session_format=session_format,
                    timezone=timezone,
                    easychair=easychair,
                    authors=authors,
                    title=title,
                    sync=sync,
                    sector=sector,
                    unesco=unesco,
                    topic=topic,
                    duration=duration,
                    start_utc=start_utc,
                    end_utc=end_utc,
                    description=description,
                )


@click.command()
@click.option("--filename", help="Custom Schedule export as xls")
def cli(filename):
    DiscourseImporter(filename)


if __name__ == "__main__":
    cli()

# import requests
# from pprint import pprint
#
# API_KEY = '3009f8fce1a0f3b1b9ccca0bcc135070557b51137c3a430db436774b98d8ea8b'
# API_HOST = 'http://localhost:3000'
#
#
# data = {
#   "title": "Imported session",
#   "raw": "<b>This is</b> a new post. Hai world.",
#   "category": 17,
# }
#
# response = requests.post("{}/conference/schedule.json".format(API_HOST), data=data, headers={
# 	"Api-Key": API_KEY,
# 	'Api-Username': 'jure'
# })
#
# try:
# 	pprint(response.json())
# except:
# 	print(response.content)
