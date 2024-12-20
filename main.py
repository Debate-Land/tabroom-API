from typing import TypedDict, List
from shared.helpers import get_tourn_boost
from pipelines.transformer import TransformedTournamentData, transform_data
from pipelines.uploader import upload_data, clear
from scraper.utils.unscraped_entries import get_unscraped_entries
from scraper.lib.entry import scrape_entry
from scraper.lib.entries import scrape_entries
from scraper.lib.tournament import scrape_tournament
from shared.helpers import enum_to_string
from pipelines.post_upload.index import update_indicies, update_all_indicies
from pipelines.post_upload.otr import update_otrs, update_all_otrs
from pipelines.post_upload.stats import update_stats, update_all_stats
from pipelines.post_upload.update_search import update_team_index, update_judge_index, update_competitor_index
from shared.lprint import lprint
from shared.const import API_BASE
import traceback
import datetime
from bullmq import Worker, Job

from scraper.lib.division import get_division_name
import requests
import requests_cache

import os
import time
import sys
import asyncio
import csv

requests_cache.install_cache("tabroom_cache")

class ScrapingJobData(TypedDict):
    class Group(TypedDict):
        id: int | None
        nickname: str

    class Season(TypedDict):
        id: int | None
        year: int

    class Division(TypedDict):
        class Circuit(TypedDict):
            id: int | None
            geographyName: str

        tabEventId: int
        event: str
        classification: str
        circuits: List[Circuit]
        firstElimRound: str | None
        tocFullBidLevel: str | None
        tournBoost: float

    group: Group
    season: Season
    tabTournId: int
    divisions: List[Division]

async def processTournament(data: ScrapingJobData, id: int | None = None):
    start = time.perf_counter()
    lprint(id, "Info", start, f"Started scraping tournament: {data['group']['nickname']} {data['season']['year']}")

    for i, division in enumerate(data['divisions']):
        lprint(id, "Info", start, f"Started scraping division {i+1}/{len(data['divisions'])}: {enum_to_string(division['classification'])} {division['event']}")

        tournament = scrape_tournament(id, data['tabTournId'])
        division_name = get_division_name(data['tabTournId'], division['tabEventId'])
        entries = []

        for entryFragment in scrape_entries(data['tabTournId'], division['tabEventId']):
            entry = scrape_entry(data['tabTournId'], entryFragment)
            entries.append(entry)

        unscraped_entries = get_unscraped_entries(entries)

        lprint(id, "Info", start, "Aggregating unscraped entries")
        while unscraped_entries:
            for tab_entry_id in unscraped_entries:
                entries.append(scrape_entry(data['tabTournId'], {
                    'code': None,
                    'location': None,
                    'school': None,
                    'tab_competitor_ids': [],
                    'tab_entry_id': tab_entry_id
                }))

            unscraped_entries = get_unscraped_entries(entries)

        lprint(id, "Info", start, message="Transforming data")
        boost = division['tournBoost']
        if boost == 1:
            boost = get_tourn_boost(division['firstElimRound'])
        data: TransformedTournamentData = transform_data(id, data['tabTournId'], division['tabEventId'], data['group']['nickname'], division['event'], tournament, entries, list(map(lambda c: c['geographyName'], division['circuits'])),
                              data['season']['year'], boost, division['classification'], division_name, enum_to_string(division['firstElimRound']), enum_to_string(division['tocFullBidLevel']), division['event'] == "PublicForum")
        lprint(id, "Info", start, message="Uploading data")
        upload_data(id, data)
        lprint(id, "Info", start, "Updating OTRs")
        update_otrs(division['tabEventId'])
        lprint(id, "Info", start, "Updating indicies")
        update_indicies(division['tabEventId'])
        lprint(id, "Info", start, "Updating stats")
        update_stats(id, division['tabEventId'])
        lprint(id, "Info", start, "Re-indexing competitors")
        update_competitor_index()
        lprint(id, "Info", start, "Re-indexing teams")
        update_team_index()
        lprint(id, "Info", start, "Re-indexing judges")
        update_judge_index()
        lprint(id, "Info", start, "Completed division")

    lprint(id, "Info", start, "Completed tournament")

async def processRetroactiveUpdate(id: int | None = None):
    start = time.perf_counter()

    lprint(id, "Info", start, f"Updating Indicies...")

    update_all_indicies(id)

    lprint(id, "Info", start, f"Updating OTRs...")

    update_all_otrs(id)

    lprint(id, "Info", start, f"Updating stats...")

    update_all_stats(id)

async def processScrapingJob(job: Job, token: str):
    try:
        await processTournament(job.data, job.id)
        await job.moveToWaitingChildren(token)
    except Exception as e:
        lprint(job.id, "Error", message=traceback.format_exc())
        raise Exception()

async def processRetroactiveUpdateJob(job: Job, token: str):
    try:
        await processRetroactiveUpdate(job.id)
        await job.moveToWaitingChildren(token)
    except Exception as e:
        lprint(job.id, "Error", message=traceback.format_exc())
        raise Exception()

async def processScrapingJobCSV(path: str):
    DATA = []

    with open(path, "r") as f:
        table = csv.reader(f)

        for row in table:
            DATA.append(row)

    for tourn in DATA:
        try:
            await processTournament({
                'group': {
                    'id': None,
                    'nickname': tourn[0]
                },
                'season': {
                    'id': None,
                    'year': int(tourn[5])
                },
                'tabTournId': int(tourn[1]),
                'divisions': [{
                    'event': tourn[3],
                    'circuits': list(map(lambda c: {
                        'id': None,
                        'geographyName': c
                    }, tourn[6].split(";"))),
                    'tabEventId': int(tourn[2]),
                    'classification': tourn[4],
                    'firstElimRound': None if tourn[7] == "None" else tourn[7],
                    'tocFullBidLevel': None if tourn[8] == "None" else tourn[8],
                    'tournBoost': float(tourn[9]),
                }]
            })
        except Exception as e:
            lprint(None, "Error", message=traceback.format_exc())
            continue

async def startWorker():
    lprint(None, "Info", message=f"Starting worker with API host `{API_BASE}`.")
    scraping_worker = Worker("scraping", processScrapingJob, { "connection": os.environ['REDIS_URL'] })
    update_worker = Worker("retroactive_update", processRetroactiveUpdateJob, { "connection": os.environ['REDIS_URL'] })

    while True:
        await asyncio.sleep(1)

if len(sys.argv) > 1 and '--file' in sys.argv and sys.argv.index('--file') > 0 and sys.argv.index('--file') < len(sys.argv) - 1:
    asyncio.run(processScrapingJobCSV(sys.argv[sys.argv.index('--file') + 1]))
elif '--retroactiveUpdate' in sys.argv:
    asyncio.run(processRetroactiveUpdate())
else:
    asyncio.run(startWorker())
