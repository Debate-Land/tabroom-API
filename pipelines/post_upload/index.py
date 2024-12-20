import time
import math
import statistics
import requests
from shared.const import API_BASE
from shared.lprint import lprint

from requests_cache import DO_NOT_CACHE, CachedSession
requests = CachedSession(expire_after=DO_NOT_CACHE)
cached_requests = CachedSession()

# Update all indicies given a tab tourn id to get judges from


def get_index_deflator(num_rounds: int, initial_index: float) -> float:
    """Gets the amount to deflate a raw index by given the number of rounds judged.

    Args:
        num_rounds (int): The number of rounds judged.
        initial_index (float): The initial index score.

    Returns:
        float: The proportion of the original Index that should be retained (from [0, 1)).
    """
    N = 0.8
    Y0 = 0.1
    K = 0.15

    rnd_deflator = N/((N/Y0 - 1)*math.pow(math.e, -
                      K*(num_rounds - 1)) + 1) + 0.9

    N = 2.2
    Y0 = 0.2
    K = 0.3

    idx_deflator = N/((N/Y0 - 1)*math.pow(math.e, -K*(initial_index - 3)) + 1)

    return rnd_deflator*idx_deflator

def update_scoped_index(judge_id: str, circuit_id: int, season_id: int) -> float:
    speaking_aggregation = cached_requests.post(f'{API_BASE}/speaking/rounds/advanced/aggregate', json={
        "where": {
            "round": {
                "result": {
                    "division": {
                        "circuits": {
                            "some": {
                                "id": circuit_id
                            }
                        },
                        "tournament": {
                            "seasonId": season_id
                        }
                    },
                }
            }
        },
        "_avg": {
            "points": True
        }
    }).json()

    scope_avg = speaking_aggregation['_avg']['points']
    if not scope_avg:
        print(f"No speaking aggregation for scope (circuit, season) = ({circuit_id}, {season_id}), defaulting to 28.5.")
        scope_avg = 28.5

    records = requests.post(f'{API_BASE}/judge-records/advanced/findMany', json={
        'where': {
            'result': {
                'judgeId': judge_id,
                'division': {
                    'tournament': {
                        'seasonId': season_id
                    },
                    'circuits': {
                        'some': {
                            'id': circuit_id
                        }
                    }
                }
            },
        },
        'include': {
            'teams': {
                'include': {
                    'rankings': {
                        'where': {
                            'seasonId': season_id,
                            'circuitId': circuit_id
                        }
                    }
                },
            },
            'rounds': {
                'include': {
                    'result': True,
                    "speaking": {
                        "where": {
                            "judgeId": judge_id
                        }
                    }
                }
            }
        }
    }).json()

    screw_sum = 0
    squirrel_sum = 0
    prelims = 0
    elims = 0
    tourns = []
    rounds = 0
    speaks = []
    pro_speaks = []
    con_speaks = []
    pro_ballots = 0
    con_ballots = 0
    low_point_wins = 0

    for record in records:
        try:
            if len(record['teams']) != 2:
                continue  # Filter out bye rounds

            if record['resultId'] not in tourns:
                tourns.append(record['resultId'])

            rounds += 1

            team_1 = record['teams'][0]['id']
            team_2 = record['teams'][1]['id']
            team_1_result = None
            team_2_result = None

            team_1_speaks = []
            team_2_speaks = []

            for round in record['rounds']:
                if round['result']['teamId'] == team_1:
                    team_1_result = round['result']
                elif round['result']['teamId'] == team_2:
                    team_2_result = round['result']

                if round['result']['teamId'] == record['winnerId']:
                    if round['side'] == "Pro":
                        pro_ballots += 1
                    elif round['side'] == "Con":
                        con_ballots += 1

                # print(round['type'])
                if round['type'] == "Prelim":
                    prelims += 1
                else:
                    elims += 1
                # print(prelims, elims)

                if round['speaking']:
                    for speak in round['speaking']:
                        points = speak['points']
                        speaks.append(points)
                        if round['side'] == "Pro":
                            pro_speaks.append(points)
                        elif round['side'] == "Con":
                            con_speaks.append(points)
                        if round['result']['teamId'] == team_1:
                            team_1_speaks.append(points)
                        elif round['result']['teamId'] == team_2:
                            team_2_speaks.append(points)

            if len(team_1_speaks) and len(team_2_speaks):
                team_1_speaks = statistics.mean(team_1_speaks)
                team_2_speaks = statistics.mean(team_2_speaks)

                if team_1_speaks > team_2_speaks and record['winnerId'] == team_2:
                    low_point_wins += 1
                elif team_2_speaks > team_1_speaks and record['winnerId'] == team_1:
                    low_point_wins += 1

            team_1_p_wp = team_1_result['prelimWins'] / (
                team_1_result['prelimWins'] + team_1_result['prelimLosses'])
            team_1_otr = record['teams'][0]['rankings'][0]['otr']
            team_1_e_wp = None

            team_2_p_wp = team_2_result['prelimWins'] / (
                team_2_result['prelimWins'] + team_2_result['prelimLosses'])
            team_2_otr = record['teams'][1]['rankings'][0]['otr']
            team_2_e_wp = None

            # Calculate & Assign EWP
            wp_high = (1.2 * abs(team_1_otr - team_2_otr))**2 / \
                statistics.mean([team_1_otr, team_2_otr]) + 0.5
            if wp_high > 0.99:
                wp_high = 0.99
            wp_low = 1 - wp_high

            if team_1_otr >= team_2_otr:
                team_1_e_wp = wp_high
                team_2_e_wp = wp_low
            else:
                team_1_e_wp = wp_low
                team_2_e_wp = wp_high

            # Record screw factor (if any) (only given in prelim rounds with incorrect result and otr delta of gte. 0.5)
            screw_factor = 0

            otr_delta = abs(team_1_otr - team_2_otr)
            avg_otr = statistics.mean([team_1_otr, team_2_otr])

            wpHi = ((1.47 * pow(otr_delta, 0.8094)) * (1 /
                    (4 * avg_otr))) / (1 + pow(2, -(10 - 6))) + 0.5
            # print("WP Hi " + str(wpHi))

            if ((record['winnerId'] == team_1 and team_1_otr < team_2_otr) or (record['winnerId'] == team_2 and team_2_otr < team_1_otr)) and record['type'] == "Prelim" and wpHi >= 0.7:
                screw_factor = 4.5**abs(0.5 - team_1_otr) - 1
                # print("Screw factor " + str(screw_factor))

            for round in record['rounds']:
                e_wp = None
                if round['result']['teamId'] == team_1:
                    e_wp = team_1_e_wp
                else:
                    e_wp = team_2_e_wp

                if not e_wp:
                    ...
                    # print(round['result']['teamId'], team_1,
                    #       team_2, team_1_e_wp, team_2_e_wp)
                    # input("Err: ")

                res = requests.post(f'{API_BASE}/rounds/advanced/update', json={
                    'where': {
                        'id': round['id']
                    },
                    'data': {
                        'expectedWinProbability': e_wp,
                    }
                })

            # Update Judge Record with screw data
            res = requests.post(f'{API_BASE}/judge-records/advanced/update', json={
                'where': {
                    'id': record['id']
                },
                'data': {
                    'screwFactor': screw_factor if round['type'] == "Prelim" else None
                }
            })

            # Update counts
            # Any nonzero screw factor counts as one screw
            screw_sum += 1 if screw_factor > 0 else 0
            squirrel_sum += 1 if record['wasSquirrel'] else 0

            # print("Screw sum: " + str(screw_sum))

        except Exception as e:
            # from pprint import pprint
            import traceback
            traceback.print_exc()
            # pprint({
            #     'where': {
            #         'result': {
            #             'judgeId': judge['id'],
            #             'division': {
            #                 'tournament': {
            #                     'seasonId': season
            #                 },
            #                 'circuits': {
            #                     'some': {
            #                         'id': circuit
            #                     }
            #                 }
            #             }
            #         },
            #     },
            #     'include': {
            #         'teams': {
            #             'include': {
            #                 'rankings': {
            #                     'where': {
            #                         'seasonId': season,
            #                         'circuitId': circuit
            #                     }
            #                 },
            #                 'results': {
            #                     'where': {
            #                         'division': {
            #                             'tabEventId': tab_event_id
            #                         }
            #                     },
            #                     'select': {
            #                         'prelimWins': True,
            #                         'prelimLosses': True
            #                     }
            #                 }
            #             },
            #         },
            #         'rounds': {
            #             'include': {
            #                 'result': True
            #             }
            #         }
            #     }
            # })
            # print(
            #     "Index computation failed on round lookup for record " + str(record['id']))
            # print(e.with_traceback(None))

    # Calculate and update index
    if not len(records):
        return print(f"No records found for judge {judge_id} on (circuit, season) = ({circuit_id}, {season_id})")
    index = 10 - (13 * squirrel_sum + 7 * screw_sum)/len(records)
    # if index < 0:
    #     index = 0
    # elif index > 10:
    #     index = 10

    # Update indicies
    print(f"{len(records)} {index}")
    index = get_index_deflator(len(records), index) * index
    print(index)

    # print(screw_sum, squirrel_sum, len(records), index)

    res = requests.post(f'{API_BASE}/rankings/judges/advanced/upsert', json={
        'where': {
            'judgeId_circuitId_seasonId': {
                'judgeId': judge_id,
                'circuitId': circuit_id,
                'seasonId': season_id
            },
        },
        'create': {
            'seasonId': season_id,
            'judgeId': judge_id,
            'circuitId': circuit_id,
            'index': index,
            'tourns': len(tourns),
            'rounds': rounds,
            'prelims': prelims,
            'elims': elims,
            'squirrels': squirrel_sum,
            'squirrelPct': squirrel_sum/len(records),
            'screws': screw_sum,
            'screwPct': screw_sum/len(records),
            'squirrelAndScrewPct': (squirrel_sum + screw_sum)/len(records),
            'avgSpks': statistics.mean(speaks) if len(speaks) else None,
            'stdSpks': statistics.stdev(speaks) if len(speaks) > 1 else None,
            'pctPro': pro_ballots/(pro_ballots + con_ballots) if pro_ballots + con_ballots != 0 else None,
            'lowPointWins': low_point_wins,
            'lowPointWinPct': low_point_wins/len(records),
            'avgProSpks': statistics.mean(pro_speaks) if len(pro_speaks) else None,
            'avgConSpks': statistics.mean(con_speaks) if len(con_speaks) else None,
            'saa': (statistics.mean(speaks) - scope_avg) if len(speaks) and scope_avg else None
        },
        'update': {
            'index': index,
            'tourns': len(tourns),
            'rounds': rounds,
            'prelims': prelims,
            'elims': elims,
            'squirrels': squirrel_sum,
            'squirrelPct': squirrel_sum/len(records),
            'screws': screw_sum,
            'screwPct': screw_sum/len(records),
            'squirrelAndScrewPct': (squirrel_sum + screw_sum)/len(records),
            'avgSpks': statistics.mean(speaks) if len(speaks) else None,
            'stdSpks': statistics.stdev(speaks) if len(speaks) > 1 else None,
            'pctPro': pro_ballots/(pro_ballots + con_ballots) if pro_ballots + con_ballots != 0 else None,
            'lowPointWins': low_point_wins,
            'lowPointWinPct': low_point_wins/len(records),
            'avgProSpks': statistics.mean(pro_speaks) if len(pro_speaks) else None,
            'avgConSpks': statistics.mean(con_speaks) if len(con_speaks) else None,
            'saa': (statistics.mean(speaks) - scope_avg) if len(speaks) and scope_avg else None
        }
    })

    if res.status_code != 200:
        lprint(None, "Judge Ranking Error", None, res.text)

def update_indicies_for_judge(judge_id: str):
    results = requests.post(f"{API_BASE}/results/judges/advanced/findMany", json={
        "where": {
            "judgeId": judge_id
        },
        "include": {
            "division": {
                "select": {
                    "circuits": {
                        "select": {
                            "id": True
                        }
                    },
                    "tournament": {
                        "select": {
                            "seasonId": True
                        }
                    }
                }
            }
        }
    }).json()

    circuits = []
    seasons = []

    for result in results:
        for circuit in result['division']['circuits']:
            circuits.append(circuit['id'])

        seasons.append(result['division']['tournament']['seasonId'])

    circuits = list(set(circuits))
    seasons = list(set(seasons))

    for season in seasons:
        for circuit in circuits:
            update_scoped_index(judge_id, circuit, season)

def update_all_indicies(job_id: int | None = None):
    judges = requests.get(f"{API_BASE}/judges").json()

    for i, judge in enumerate(judges):
        if i % 500 == 0:
            lprint(job_id, "Info", message=f"Updating {i+1}/{len(judges)}")
        update_indicies_for_judge(judge['id'])

def update_indicies(tab_event_id: int):
    """_summary_

    Args:
        tab_event_id (int): _description_
    """

    # TODO: Abstract
    event = requests.post(f'{API_BASE}/tournaments/divisions/advanced/findUnique', json={
        'where': {
            'tabEventId': tab_event_id
        },
        'select': {
            'teamResults': {
                'select': {
                    'teamId': True
                }
            },
            'circuits': {
                'select': {
                    'id': True
                }
            },
            'tournament': {
                'select': {
                    'seasonId': True
                }
            }
        }
    }).json()

    circuits = list(map(lambda c: c['id'], event['circuits']))
    season = event['tournament']['seasonId']

    # circuits = [34, 25, 26, 27, 28, 29, 30, 31, 32, 33]
    # season = 16
    # judges = requests.get(f"{API_BASE}/judges").json()
    judges = requests.post(f'{API_BASE}/judges/advanced/findMany', json={
        'where': {
            'results': {
                'some': {
                    'division': {
                        'tabEventId': tab_event_id
                    }
                }
            }
        }
    }).json()

    print(f"Found {len(judges)} who judged at {tab_event_id}")

    for circuit in circuits:
        print(f"Circuit {circuit}")

        for i, judge in enumerate(judges):
            print(f"{i+1}/{len(judges)} {judge['id']}")
            update_scoped_index(judge['id'], circuit, season)

if __name__ == "__main__":
    # tourns = requests.get(
    #     f"{API_BASE}/tournaments/divisions").json()
    # for tourn in tourns:
    #     print(f"Updating {tourn['tabEventId']}")
    #     update_indicies(tourn['tabEventId'])
    # print(get_index_deflator(20, 7.5))
    update_indicies(242828)
