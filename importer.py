import click
import requests
import xlrd
import arrow
from pprint import pprint
from jinja2 import Template, environment
import time
from env import API_KEY, API_USER, API_HOST

CATEGORIES = {
    "Applications of Open Education Practices/Open Pedagogy/Open Education Research": "Practices, Pedagogies & Research",
    "Connecting Open Education to Primary and Secondary (K-12) Education": "Primary Secondary",
    "Global Collaboration, Strategies, & Policies in Open Education": "Policies",
    "Innovation through MOOCs practices": "MOOCs",
    "Technologies for Open Education": "OE Technologies",
}

SYNC_ID = 17
ASYNC_ID = 25


def pluralize(number, singular="", plural="s"):
    if number == 1:
        return singular
    else:
        return plural


class DiscourseImporter(object):
    def _new_post(self, **kwargs):
        topic = kwargs.get("topic")
        if topic:
            body = self.post_template.render(**kwargs)

            if kwargs.get("sync") == "sync":
                sync = True
                topic_id = SYNC_ID
                title = "🎤 {}".format(kwargs.get("title"))
            else:
                sync = False
                topic_id = ASYNC_ID
                title = "📽️ {}".format(kwargs.get("title"))

            data = {
                "title": title,
                "raw": body,
                "category": topic_id,
                "tags[]": [
                    "oeg20_{}".format(kwargs["easychair"]),
                    CATEGORIES[topic],
                    kwargs.get("session_format"),
                ],
            }

            response = requests.post(
                "{}/posts.json".format(API_HOST),
                data=data,
                headers={"Api-Key": API_KEY, "Api-Username": API_USER},
            ).json()

            print(response)
            url = "/t/{}/{}".format(response["topic_slug"], response["topic_id"])

            data = {
                "title": kwargs.get("title"),
                "start": kwargs.get("start_utc"),
                "end": kwargs.get("end_utc"),
                "url": url,
                "topic": kwargs.get("topic"),
                "format": kwargs.get("session_format"),
                "author": kwargs.get("authors"),
                "sync": sync,
                "easychair": kwargs.get("easychair"),
                "timezone": kwargs.get("timezone"),
                "unesco": kwargs.get("unesco"),
            }

            response = requests.post(
                "{}/conference/schedule.json".format(API_HOST),
                data=data,
                headers={"Api-Key": API_KEY, "Api-Username": API_USER},
            ).json()

            pprint(response)
            time.sleep(1)

    def __init__(self, filename):
        self.workbook = xlrd.open_workbook(filename)
        with open("templates/post.html") as file_:
            environment.DEFAULT_FILTERS["pluralize"] = pluralize
            self.post_template = Template(file_.read())

        self.clear_schedule()
        self.create_topics()

    def clear_schedule(self):
        response = requests.get(
            "{}/conference/schedule.json".format(API_HOST),
            headers={"Api-Key": API_KEY, "Api-Username": API_USER},
        ).json()
        print(response)
        print(len(response["conference_plugin"][0]["conference_plugin"]))

        response = requests.delete(
            "{}/conference/clear.json".format(API_HOST),
            headers={"Api-Key": API_KEY, "Api-Username": API_USER},
        )
        print(response)

        response = requests.get(
            "{}/conference/schedule.json".format(API_HOST),
            headers={"Api-Key": API_KEY, "Api-Username": API_USER},
        ).json()
        print(response)

    def create_topics(self):
        sheet = self.workbook.sheet_by_name("Easychair Export")
        data = {}
        for row_idx in range(2, sheet.nrows):
            try:
                easychair = int(sheet.cell(row_idx, 0).value)
            except ValueError:
                continue

            title = sheet.cell(row_idx, 1).value
            keywords = sheet.cell(row_idx, 4).value
            abstract = sheet.cell(row_idx, 6).value

            data[easychair] = {
                "title": title,
                "keywords": keywords.split("\n"),
                "abstract": abstract,
            }

        sheet = self.workbook.sheet_by_name("Authors")
        authors = {}
        for row_idx in range(2, sheet.nrows):
            try:
                easychair = int(sheet.cell(row_idx, 0).value)
            except ValueError:
                continue

            name = "{} {}".format(
                sheet.cell(row_idx, 1).value, sheet.cell(row_idx, 2).value
            )
            country = sheet.cell(row_idx, 4).value
            org = sheet.cell(row_idx, 5).value

            if not authors.get(easychair):
                authors[easychair] = {"names": [], "countries": [], "orgs": []}

            authors[easychair]["names"].append(name)
            authors[easychair]["countries"].append(country)
            authors[easychair]["orgs"].append(org)

        for sheetname, tz in [
            ("Taiwan-16-18-20", "Asia/Taipei"),
            ("Netherlands-16-18-20", "Europe/Berlin"),
            ("Canada-16-18-20", "America/Toronto"),
            ("Async & NA", None),
        ]:
            sheet = self.workbook.sheet_by_name(sheetname)
            for row_idx in range(
                2,
                # sheet.nrows
                5,
            ):
                session_format = sheet.cell(row_idx, 0).value
                timezone = sheet.cell(row_idx, 1).value
                try:
                    easychair = int(sheet.cell(row_idx, 2).value)
                except ValueError:
                    continue
                # authors = sheet.cell(row_idx, 3).value
                title = sheet.cell(row_idx, 4).value
                sync = sheet.cell(row_idx, 5).value
                sector = sheet.cell(row_idx, 6).value
                unesco = sheet.cell(row_idx, 7).value
                topic = sheet.cell(row_idx, 8).value
                try:
                    description = sheet.cell(row_idx, 13).value
                except IndexError:
                    description = None

                if tz:
                    duration = sheet.cell(row_idx, 9).value
                    duration = xlrd.xldate_as_tuple(duration, self.workbook.datemode)
                    date = sheet.cell(row_idx, 10).value
                    date = xlrd.xldate_as_tuple(date, self.workbook.datemode)
                    start = sheet.cell(row_idx, 11).value
                    start = xlrd.xldate_as_tuple(start, self.workbook.datemode)
                    end = sheet.cell(row_idx, 12).value
                    end = xlrd.xldate_as_tuple(end, self.workbook.datemode)
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
                else:
                    duration = None
                    start_utc = None
                    end_utc = None

                self._new_post(
                    session_format=session_format,
                    timezone=timezone,
                    easychair=easychair,
                    authors=authors.get(easychair, {}).get("names"),
                    title=title,
                    sync=sync,
                    sector=sector,
                    unesco=unesco,
                    topic=topic,
                    duration=duration,
                    start_utc=start_utc,
                    end_utc=end_utc,
                    description=data.get(easychair, {}).get("abstract"),
                    keywords=data.get(easychair, {}).get("keywords", []),
                    orgs=authors.get(easychair, {}).get("orgs"),
                    countries=authors.get(easychair, {}).get("countries"),
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
