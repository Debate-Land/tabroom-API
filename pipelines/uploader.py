import requests
from shared.lprint import lprint
from shared.const import API_BASE
from typing import Mapping
from .transformer import TransformedTournamentData
from requests_cache import DO_NOT_CACHE, CachedSession
from random import random
from tests.runtime import process_runtime_tests
requests = CachedSession(expire_after=DO_NOT_CACHE)


def clear():
    TABLES = [
        # 'speaking/rounds',
        # 'speaking/tournaments',
        # 'bids',
        # 'competitors',
        # 'rankings/teams',
        # 'rankings/judges',
        # 'judge-records',
        # 'rounds',
        # 'results/judges',
        # 'results/teams',
        # 'aliases',
        # 'teams',
        # 'paradigms/links',
        # 'paradigms/emails',
        # 'paradigms',
        'judge-notes',
        'strikesheets',
        'judges',
        'tournaments/fragments',
        'tabroom-circuits',
        'tournaments/assets',
        'tournaments/pages',
        'tournaments/sites',
        'tabroom-emails',
        'tournaments/event-metadata',
        'schools',
        'tournaments/divisions',
        'tournaments',
        'tournaments/groups',
        'circuits',
        'seasons',

    ]

    for TABLE in TABLES:
        print("Deleting " + TABLE)
        res = requests.delete(f'{API_BASE}/{TABLE}')

        if res.status_code != 200:
            print("Error.")
            quit()


def upload_data(job_id: int | None, data: TransformedTournamentData):
    # Process Tournament
    tournament = data['tournament']

    tournament_body = {
        'tabTournId': tournament['tab_tourn_id'],
        'season': {
            'connectOrCreate': {
                'where': {
                    'year': tournament['season']
                },
                'create': {
                    'year': tournament['season']
                }
            }
        },
        'registrationOpens': tournament['registration_opens'],
        'registrationCloses': tournament['registration_closes'],
        'feesFrozen': tournament['fees_frozen'],
        'webname': tournament['webname'],
        'contacts': {
            'connectOrCreate': list(map(
                lambda c: {
                    'where': {
                        'email': c['email']
                    },
                    'create': {
                        'email': c['email'],
                        'name': c['name']
                    }
                },
                tournament['contacts']
            ))
        },
        'pages': {
            'createMany': {
                'data': list(map(lambda p: {
                    'title': p['title'],
                    'tabWebpageId': p['tab_webpage_id'],
                    'html': p['html'],
                    'text': p['text']
                }, tournament['pages']))
            }
        },
        'assets': {
            'createMany': {
                'data': tournament['assets']
            }
        },
        'tabCircuits': {
            'connectOrCreate': list(map(
                lambda c: {
                    'where': {
                        'tabCircuitId': c['tab_circuit_id']
                    },
                    'create': {
                        'tabCircuitId': c['tab_circuit_id'],
                        'abbreviation': c['abbreviation'],
                        'name': c['name']
                    }
                },
                tournament['tab_circuits']
            ))
        },
        'sites': {
            'connectOrCreate': list(map(
                lambda s: {
                    'where': {
                        'tabSiteId': s['tab_site_id']
                    },
                    'create': {
                        'name': s['name'],
                        'host': s['host'],
                        'tabSiteId': s['tab_site_id']
                    }
                },
                tournament['sites']
            ))
        },
        'pastResults': {
            'connectOrCreate': list(map(
                lambda r: {
                    'where': {
                        'tabTournId': r['tab_tourn_id']
                    },
                    'create': {
                        'name': r['name'],
                        'tabTournId': r['tab_tourn_id'],
                        'season': {
                            'connectOrCreate': {
                                'where': {
                                    'year': r['year']
                                },
                                'create': {
                                    'year': r['year']
                                }
                            }
                        }
                    }
                },
                tournament['past_results']
            ))
        },
        'eventMetadatum': {
            'createMany': {
                'data': list(map(
                    lambda d: {
                        'topic': d['topic'],
                        'topicClassification': d['topic_classification'],
                        'abbreviation': d['abbreviation'],
                        'format': d['format'],
                        'entryFee': d['entry_fee'],
                        'schoolEntryLimit': d['school_entry_limit'],
                        'competitorsPerEntry': ' '.join(map(lambda e: str(e), sorted(d['competitors_per_entry'])))
                    },
                    tournament['event_metadata']))
            }
        },
        'emails': {
            'createMany': {
                'data': tournament['emails']
            }
        },
        'group': {
            'connectOrCreate': {
                'where': {
                    'nickname': tournament['nickname']
                },
                'create': {
                    'nickname': tournament['nickname']
                }
            }
        }
    }
    tournament_body_create = tournament_body
    tournament_body_create['name'] = tournament['name']
    tournament_body_create['location'] = tournament['location']
    tournament_body_create['start'] = tournament['start']
    tournament_body_create['end'] = tournament['end']

    tournament_res = requests.post(f'{API_BASE}/tournaments/advanced/upsert', json={
        'where': {
            'group': {
                'nickname': tournament['nickname']
            },
            'season': {
                'year': tournament['year']
            }
        },
        'create': tournament_body_create,
        'update': tournament_body
    })

    if tournament_res.status_code != 200:
        message = f"Could not upsert tournament. {tournament_res.text}"
        lprint(job_id, "Error", message=message)
        # raise TypeError("Invalid API Response. " + message)

    geography_2_id = {}

    for geography in tournament['circuits']:
        res = requests.post(f'{API_BASE}/geographies/advanced/upsert', json={
            'where': {
                'name': geography
            },
            'create': {
                'name': geography,
                'abbreviation': geography
            },
            'update': {}
        }).json()

        geography_2_id[geography] = res['id']

    # Process Tournament Event
    tournament_division_body = {
        'tabEventId': tournament['tab_event_id'],
        'circuits': {
            'connectOrCreate': list(map(
                lambda c: {
                    'where': {
                        'geographyId_event_classification': {
                            'event': tournament['event'],
                            'geographyId': geography_2_id[c],
                            'classification': tournament['classification']
                        }
                    },
                    'create': {
                        'geography': {
                            'connect': {
                                'id': geography_2_id[c]
                            }
                        },
                        'event': tournament['event'],
                        'classification': tournament['classification'],
                        'seasons': {
                            'connectOrCreate': {
                                'where': {
                                    'year': tournament['season']
                                },
                                'create': {
                                    'year': tournament['season']
                                }
                            }
                        }
                    }
                },
                tournament['circuits']
            ))
        },
        'event': tournament['event'],
        'name': tournament['division_name'],
        'firstElimRound': "".join(tournament['first_elim_round'].split(" ")) if tournament['first_elim_round'] else None,
        'bidLevel': tournament['bid_level'],
        'classification': tournament['classification'],
        'tournamentId': tournament_res.json()['id'],
    }

    tournament_division_res = requests.post(f'{API_BASE}/tournaments/divisions/advanced/upsert', json={
        'where': {
            'tabEventId': tournament['tab_event_id']
        },
        'create': tournament_division_body,
        'update': tournament_division_body
    })

    if tournament_division_res.status_code != 200:
        message = f"Could not upsert tournament division. {tournament_division_res.text}"
        lprint(job_id, "Error", message=message)
        # raise TypeError("Invalid API Response. " + message)

    tournament_division_id = tournament_division_res.json()['id']
    competing_school_ids = []

    team_id_to_result_id: Mapping[str, int] = {}
    judge_id_to_result_id: Mapping[str, int] = {}

    team_upload_messages = []

    # Create teams & all schools/aliases
    for result in data['team_results']:
        # Create team and upsert and/or connect school and alias
        team_body = {
            'id': result['team_id'],
            'competitors': {
                'connectOrCreate': list(map(
                    lambda c: {
                        'where': {
                            'id': c['id']
                        },
                        'create': c
                    },
                    result['competitors']
                ))
            },
            'schools': {
                'connectOrCreate': {
                    'where': {
                        'name': result['school']
                    },
                    'create': {
                        'name': result['school'],
                        # TODO: Add country support
                        'state': result['location']['state'] if result['location'] else None
                    }
                }
            },
            'aliases': {
                'connectOrCreate': {
                    'where': {
                        'code_teamId': {
                            'code': result['code'],
                            'teamId': result['team_id']
                        }
                    },
                    'create': {
                        'code': result['code']
                    }
                }
            }
        }
        team_res = requests.post(f'{API_BASE}/teams/advanced/upsert', json={
            'where': {
                'id': result['team_id']
            },
            'create': team_body,
            'update': team_body,
            'select': {
                'aliases': {
                    'where': {
                        'code': result['code']
                    },
                    'select': {
                        'id': True
                    }
                },
                'schools': {
                    'where': {
                        'name': result['school']
                    },
                    'select': {
                        'id': True
                    }
                }
            }
        })

        school_id = requests.get(
            f'{API_BASE}/schools?where="name":"{result["school"]}"').json()[0]['id']
        # alias_id = requests.get(
        #     f'{API_BASE}/aliases?where="code":"{result["code"]}","teamId":"{result["team_id"]}"').json()[0]['id']
        alias_id = requests.post(f'{API_BASE}/aliases/advanced/findFirst', data={
            'where': {
                'code': result['code'],
                'teamId': result['team_id']
            }
        }).json()['id']

        # Create team result
        result_body = {
            'divisionId': tournament_division_id,
            'teamId': result['team_id'],
            'tabEntryId': result['tab_entry_id'],
            'aliasId': alias_id,
            'schoolId': school_id,
            'prelimPos': result['prelim_pos'],
            'prelimPoolSize': result['prelim_pool_size'],
            'prelimWins': result['prelim_wins'],
            'prelimLosses': result['prelim_losses'],
            'prelimBallotsWon': result['prelim_ballots_won'],
            'prelimBallotsLost': result['prelim_ballots_lost'],
            'elimWins': result['elim_wins'],
            'elimLosses': result['elim_losses'],
            'elimBallotsWon': result['elim_ballots_won'],
            'elimBallotsLost': result['elim_ballots_lost'],
            'opWpM': result['op_wp_m'],
            'otrComp': result['otr_comp'],
            'speaking': {
                'create': list(map(
                    lambda speak: {
                        'competitorId': speak['competitor_id'],
                        'rawAvgPoints': speak['raw_avg_points'],
                        'adjAvgPoints': speak['adj_avg_points'],
                        'stdDevPoints': speak['std_dev_points']
                    },
                    result['speaking']
                ))
            },
        }

        if result['bid']:
            result_body['bid'] = {
                'create': {
                    'value': result['bid']['value'],
                    'isGhostBid': result['bid']['is_ghost_bid'] if 'is_ghost_bid' in result['bid'] else False,
                    'team': {
                        'connect': {
                            'id': result['team_id']
                        }
                    },
                }
            }

        result_res = requests.post(
            f'{API_BASE}/results/teams', json=result_body)

        if result_res.status_code != 200:
            team_upload_messages.append(
                f"Could not upsert team result. {result_res.text}")
            # raise TypeError("Invalid API Response. " + message)
        else:
            team_id_to_result_id[result['team_id']] = result_res.json()['id']

        # Add schools to list if not already in
        if school_id not in competing_school_ids:
            competing_school_ids.append(school_id)

    if len(team_upload_messages):
        lprint(job_id, "Error", message="\n".join(team_upload_messages))

    # TODO: connect all schools to the tournament event
    for school_id in competing_school_ids:
        requests.patch(f"{API_BASE}/schools/{school_id}", json={
            "divisions": {
                "connect": {
                    "tabEventId": tournament_division_body['tabEventId']
                }
            }
        })

    #  Create all judges and their results
    for result in data['judge_results']:
        if result['judge_id'] in judge_id_to_result_id:
            lprint(job_id, "Warning",
                   message=f"Already scraped {result['judge_id']}")
            continue

        judge_body = {
            'id': result['judge_id'],
            'name': result['name']
        }

        # TODO: School affiliation
        judge_res = requests.post(f'{API_BASE}/judges/advanced/upsert', json={
            'where': {
                'id': result['judge_id']
            },
            'create': judge_body,
            'update': {}
        })

        result_body = {
            'judgeId': result['judge_id'],
            'tabJudgeId': result['tab_judge_id'],
            'avgRawPoints': result['avg_raw_points'],
            'avgAdjPoints': result['avg_adj_points'],
            'stdDevPoints': result['std_dev_points'],
            'numPrelims': result['num_prelims'],
            'numScrews': result['num_screws'],
            'numElims': result['num_elims'],
            'numSquirrels': result['num_squirrels'],
            'numPro': result['num_pro'],
            'numCon': result['num_con'],
            'divisionId': tournament_division_id
        }

        result_res = requests.post(
            f'{API_BASE}/results/judges', json=result_body)
        try:
            judge_id_to_result_id[result['judge_id']] = result_res.json()['id']
        except Exception:
            ...  # Merge results a second time by tabJudgeId in transformer

    # Create rounds
    for _round in data['rounds']:
        if _round['team_id'] not in team_id_to_result_id:
            continue

        round_body = {
            'name': _round['name'],
            'nameStd': _round['name_std'],
            'type': _round['type'],
            'side': _round['side'],
            'outcome': _round['outcome'],
            'ballotsWon': _round['ballots_won'],
            'ballotsLost': _round['ballots_lost'],
            'opponentId': _round['opponent_id'],
            'resultId': team_id_to_result_id[_round['team_id']],
            'speaking': {
                'create': _round['speaking']
            }
        }

        round_res = requests.post(f'{API_BASE}/rounds/advanced/create', json={
            "data": round_body
        })
        if round_res.status_code != 200:
            print(round_body)
            # input("Continue: ")

    # Create records
    for record in data['records']:
        if record['judge_id'] not in judge_id_to_result_id:
            continue

        record_body = {
            'decision': record['decision'],
            'avgPoints': record['avg_points'],
            'wasSquirrel': record['was_squirrel'],
            'judgeId': record['judge_id'],
            'teams': {
                'connect': list(map(
                    lambda id: {'id': id},
                    record['teams']
                ))
            },
            'winnerId': record['winner_id'],
            'type': record['type'],
            'event': data['tournament']['event'],
            'resultId': judge_id_to_result_id[record['judge_id']],
            'rounds': {
                'connect': list(map(
                    lambda team_id: {
                        'nameStd_resultId': {
                            'nameStd': record['round_name_std'],
                            'resultId': team_id_to_result_id[team_id]
                        }
                    },
                    record['teams']
                ))
            }
        }

        record_res = requests.post(
            f'{API_BASE}/judge-records', json=record_body)
        if record_res.status_code != 200:
            # input("Err: ")
            ...

    # Create paradigms
    for paradigm in data['paradigms']:
        paradigm_body = {
            'id': paradigm['hash'],
            'scrapedAt': paradigm['scraped_at'],
            'text': paradigm['text'],
            'html': paradigm['html'],
            'judge': {
                'connect': {
                    'id': paradigm['judge_id']
                }
            },
            'flowRating': paradigm['flow_confidence'],
            'progressiveRating': paradigm['progressive_confidence'],
            'emails': {
                'connectOrCreate': list(map(lambda e: {
                    'create': {
                        'email': e
                    },
                    'where': {
                        'email': e
                    }
                }, paradigm['emails']))
            },
            'links': {
                'connectOrCreate': list(map(lambda l: {
                    'create': {
                        'href': l
                    },
                    'where': {
                        'href': l
                    }
                }, paradigm['links']))
            }
        }

        paradigm_res = requests.post(
            f'{API_BASE}/paradigms', json=paradigm_body)

    # Runtime tests
    # process_runtime_tests(job_id, data)
