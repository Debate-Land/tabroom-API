import math
import time
import requests
from shared.const import API_BASE
from shared.lprint import lprint
from requests_cache import DO_NOT_CACHE, CachedSession
requests = CachedSession(expire_after=DO_NOT_CACHE)

def get_otr_deflator(numTourns: int) -> float:
    """Gets the amount to deflate a raw OTR average by given the number of tournaments attended.

    Args:
        numTourns (int): The number of tournaments attended.

    Returns:
        float: The proportion of the original OTR that should be retained (from [0, 1)).
    """
    N = 1
    Y0 = 0.04  # g
    K = 1.3

    return round(N/((N/Y0 - 1)*math.pow(math.e, -K*numTourns) + 1), 2)

# def get_break_boost(numElimRoundsWon: int) -> float:
#     F = 5
#     D = 0.8002
#     C = 3.2
#     B = 0.7

#     return round(F/(1+math.pow(math.e, -B*numElimRoundsWon + C)) + D, 2)

def update_scoped_otr(team_id: str, circuit_id: int, season_id: int):
    results = requests.post(f'{API_BASE}/results/teams/advanced/findMany', json={
        'where': {
            'teamId': team_id,
            'division': {
                'circuits': {
                    'some': {
                        'id': circuit_id
                    }
                },
                'tournament': {
                    'seasonId': season_id
                }
            },
        },
        'select': {
            'otrComp': True,
            'id': True
        }
    }).json()

    if not len(results):
        return print(f"No results for team {team_id} on (circuit, season) = ({circuit_id}, {season_id})")

    otrs = list(map(lambda r: r['otrComp'], results))
    updated_otr = get_otr_deflator(len(otrs)) * sum(otrs)/len(otrs)
    rankings_res = requests.post(f'{API_BASE}/rankings/teams/advanced/upsert', json={
        'where': {
            'teamId_circuitId_seasonId': {
                'teamId': team_id,
                'seasonId': season_id,
                'circuitId': circuit_id
            }
        },
        'create': {
            'team': {
                'connect': {
                    'id': team_id
                }
            },
            'season': {
                'connect': {
                    'id': season_id
                }
            },
            'circuit': {
                'connect': {
                    'id': circuit_id
                }
            },
            'otr': updated_otr
        },
        'update': {
            'otr': updated_otr
        }
    })

def update_otrs_for_team(team_id: str):
    results = requests.post(f"{API_BASE}/results/teams/advanced/findMany", json={
        "where": {
            "teamId": team_id
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
            print("Updating", circuit, season)
            update_scoped_otr(team_id, circuit, season)

def update_all_otrs(job_id: int | None = None):
    teams = requests.get(f"{API_BASE}/teams").json()

    for i, team in enumerate(teams):
        lprint(job_id if i % 500 == 0 else None, "Info", message=f"Updating {i+1}/{len(teams)}")
        update_otrs_for_team(team['id'])
        time.sleep(0.5)

def update_otrs(tab_event_id: int) -> None:
    """Updates all OTRs for all entries in an event (tab_event_id).

    Args:
        tab_event_id (int): The unique ID assigned to the event, visible in the `event_id` query parameter.
    """

    # Get all teamIds participating and the season + all circuits in scope
    # lprint("Getting teams")
    # teams = list(map(lambda t: t['id'], requests.get(
    #     f"{API_BASE}/teams").json()))

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
    # circuits = [25, 26, 27, 28, 29, 30, 31, 32, 33, 34]
    # season = 16
    teams = list(map(lambda r: r['teamId'], event['teamResults']))

    for circuit in circuits:
        print(f"Circuit {circuit}")
        for i, team in enumerate(teams):
            print(f"{i+1}/{len(teams)} {team}")
            update_scoped_otr(team, circuit, season)

if __name__ == "__main__":
    update_otrs_for_team("dfd0b582a4ef75895f07cb21")
    # update_otrs(242828)
