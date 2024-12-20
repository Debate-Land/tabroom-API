import statistics
import requests
import time
from typing import List
from shared.const import API_BASE
from shared.lprint import lprint
from requests_cache import DO_NOT_CACHE, CachedSession
requests = CachedSession(expire_after=DO_NOT_CACHE)

def _hi_lo_avg(speaks: List[float], trim: int) -> float | None:
    """Removes the high/low speaks from the given list after sorting
    and returns the average.

    Args:
        speaks (List[float]): A list of speaker points.
        trim (int): The number to remove from each end.

    Returns:
        float: The trimmed average, or None if there's a fallback.
    """

    # Sort the speaks list
    sorted_speaks = sorted(speaks)

    # Check if the list has enough elements to trim
    if len(sorted_speaks) <= 2 * trim:
        return None

    # Trim the highest and lowest values
    trimmed_speaks = sorted_speaks[trim:-trim]

    # Calculate the average of the remaining values
    if trimmed_speaks:  # Check if the list is not empty
        return sum(trimmed_speaks) / len(trimmed_speaks)
    else:
        return None


def _update_scoped_stats(job_id: int | None, season: int, circuit: int):
    """_summary_

    Args:
        tournament_division_id (int): _description_
    """

    # Get all the team results in the scope
    team_results = requests.post(f'{API_BASE}/results/teams/advanced/findMany', json={
        'where': {
            'division': {
                'circuits': {
                    'some': {
                        'id': circuit
                    }
                },
                'tournament': {
                    'seasonId': season
                }
            },
        },
        'include': {
            'rounds': {
                'include': {
                    'records': {
                        'include': {
                            'judge': {
                                'include': {
                                    'rankings': {
                                        'where': {
                                            'seasonId': season,
                                            'circuitId': circuit
                                        }
                                    },
                                    'paradigms': {
                                        'select': {
                                            'flowRating': True,
                                            'progressiveRating': True
                                        },
                                        'orderBy': {
                                            'scrapedAt': 'desc',
                                        },
                                        'take': 1
                                    }
                                }
                            },
                            'rounds': {
                                'include': {
                                    'speaking': True
                                }
                            }
                        }
                    },
                    'speaking': True,
                    'result': {
                        'select': {
                            'id': True
                        }
                    }
                }
            },
            'team': {
                'include': {
                    'rankings': {
                        'where': {
                            'seasonId': season,
                            'circuitId': circuit
                        }
                    }
                }
            },
            'bid': True
        }
    }).json()

    # Maps
    team_2_rounds = {}
    team_2_bids = {}
    team_2_results = {}
    team_2_otr = {}
    team_2_speaks = {}
    judge_2_speaks = {}
    judge_2_index = {}
    speaks = []

    # Iterate over all team results
    for result in team_results:
        # Upsert result into map
        if result['teamId'] not in team_2_results:
            team_2_results[result['teamId']] = []

        team_2_results[result['teamId']].append(result['id'])

        # Upsert rounds into map
        if result['teamId'] not in team_2_rounds:
            team_2_rounds[result['teamId']] = []

        [team_2_rounds[result['teamId']].append(
            round) for round in result['rounds']]

        # Add bids
        if result['teamId'] not in team_2_bids:
            team_2_bids[result['teamId']] = 0
        if result['bid']:
            team_2_bids[result['teamId']
                        ] += 1 if result['bid']['value'] == "Full" else 0.5

        # Upsert result into map
        if result["teamId"] not in team_2_rounds:
            team_2_rounds[result['teamId']] = []

        team_2_rounds[result['teamId']].append(result['id'])

        # Set OTR in map
        if not len(result['team']['rankings']):
            continue
        team_2_otr[result['teamId']] = result['team']['rankings'][0]['otr']

        # Upsert speaks into map
        if result['teamId'] not in team_2_speaks:
            team_2_speaks[result['teamId']] = []

        for round in result['rounds']:
            for record in round['records']:
                # Set index in map
                if len(record['judge']['rankings']):
                    judge_2_index[record['judgeId']
                                  ] = record['judge']['rankings'][0]['index']
                else:
                    lprint(None, event="Warning", message=f"No ranking found for judge {record['judgeId']} on season {season} and circuit {circuit}")

                # Upsert judge speaks into map
                if record['judgeId'] not in judge_2_speaks:
                    judge_2_speaks[record['judgeId']] = []
                if 'speaking' in round:
                    for speak in round['speaking']:
                        judge_2_speaks[record['judgeId']].append(
                            speak['points'])
                        team_2_speaks[result['teamId']].append(speak['points'])
                        speaks.append(speak['points'])

    avg_speaks = statistics.mean(speaks) if len(speaks) else None

    # with open("foo.json", "w") as f:
    #     import json
    #     json.dump({
    #         'team_2_otr': team_2_otr,
    #         'team_2_bids': team_2_bids,
    #         'team_2_results': team_2_results,
    #         'team_2_speaks': team_2_speaks,
    #         'judge_2_speaks': judge_2_speaks,
    #         'judge_2_index': judge_2_speaks,
    #     }, f)
    for i, (teamId, rounds) in enumerate(team_2_rounds.items()):
        lprint(job_id if i % 500 == 0 else None, "Info", message=f"Updating {i+1}/{len(team_2_rounds.items())}")
        x_wp = []
        # We don't have an expWp for ALL rounds (eg. bye)
        wins_with_exp_wp_recorded = 0
        # We don't have an expWp for ALL rounds (eg. bye)
        losses_with_exp_wp_recorded = 0
        prelim_wins = 0
        prelim_losses = 0
        prelim_splits = 0
        elim_wins = 0
        elim_losses = 0
        underdog_wins = 0
        underdog_losses = 0
        favorite_wins = 0
        favorite_losses = 0
        high_point_wins = 0
        high_point_losses = 0
        low_point_wins = 0
        low_point_losses = 0
        opp_speaker_impact = []
        speaks_above_judging = []
        speaks_above_average = []
        competitor_2_speaks = {}
        avg_team_speaks = []
        tourn_to_speaks = {}
        speaks_1hl = []
        speaks_2hl = []
        tourns = len(team_2_results[teamId])
        num_rounds = len(
            list(filter(lambda r: not isinstance(r, int) and r['outcome'] != "Split", team_2_rounds[result['teamId']])))
        lay_record = [0, 0]
        flow_record = [0, 0]
        trad_record = [0, 0]
        prog_record = [0, 0]
        wins_above_replacement = None
        win_pct_above_replacement = None
        true_wp = None
        intra_team_speak_std = None
        d_sp = None
        overall_speak_std = statistics.stdev(
            team_2_speaks[teamId]) if teamId in team_2_speaks and len(team_2_speaks[teamId]) >= 2 else None

        for round in rounds:
            if isinstance(round, int):
                continue

            # xWp and related stats
            if round['expectedWinProbability'] != None:
                x_wp.append(round['expectedWinProbability'])
                if round['outcome'] == "Win":
                    wins_with_exp_wp_recorded += 1
                if round['expectedWinProbability'] < 0.5:
                    if round['outcome'] == "Win":
                        underdog_wins += 1
                    elif round['outcome'] == "Loss":
                        underdog_losses += 1
                elif round['expectedWinProbability'] > 0.5:
                    if round['outcome'] == "Win":
                        favorite_wins += 1
                    elif round['outcome'] == "Loss":
                        favorite_losses += 1

            # Paradigms stats
            for record in round['records']:
                paradigms = record['judge']['paradigms']

                if len(paradigms):
                    flow = paradigms[0]['flowRating']
                    prog = paradigms[0]['progressiveRating']

                    wasWinner = record['winnerId'] == teamId

                    # Flow/lay
                    if flow < 6:
                        if wasWinner:
                            lay_record[0] += 1
                        else:
                            lay_record[1] += 1
                    else:
                        if wasWinner:
                            flow_record[0] += 1
                        else:
                            flow_record[1] += 1

                    # Prog/trad
                    if prog < 4:
                        if wasWinner:
                            trad_record[0] += 1
                        else:
                            trad_record[1] += 1
                    else:
                        if wasWinner:
                            prog_record[0] += 1
                        else:
                            prog_record[1] += 1

            # Prelim/Elim stats
            if round['type'] == "Prelim":
                if round['outcome'] == "Win":
                    prelim_wins += 1
                elif round['outcome'] == "Loss":
                    prelim_losses += 1
                else:
                    prelim_splits += 1

            else:
                if round['outcome'] == "Win":
                    elim_wins += 1
                elif round['outcome'] == "Loss":
                    elim_losses += 1

            if round['speaking']:
                team_round_avg = []
                judge_2_team_avg_points = {}

                # Handle competitor point lookup and judge avg points
                for speak in round['speaking']:
                    competitor = speak['competitorId']
                    judge = speak['judgeId']
                    points = speak['points']

                    team_round_avg.append(points)

                    if competitor not in competitor_2_speaks:
                        competitor_2_speaks[competitor] = []
                    competitor_2_speaks[competitor].append(points)

                    if judge not in judge_2_team_avg_points:
                        judge_2_team_avg_points[judge] = []
                    judge_2_team_avg_points[judge].append(points)

                for judgeId, points in judge_2_team_avg_points.items():
                    judge_2_team_avg_points[judgeId] = statistics.mean(points)

                # Handle opponent speaker avg
                opponent_round_avg = []
                for record in round['records']:
                    for subround in record['rounds']:
                        if subround['opponentId'] != teamId:
                            continue
                        if 'speaking' in subround:
                            for speak in subround['speaking']:
                                opponent_round_avg.append(speak['points'])

                # Record win <-> speaking stats
                team_round_avg = statistics.mean(
                    team_round_avg) if len(team_round_avg) else 0
                opponent_round_avg = statistics.mean(
                    opponent_round_avg) if len(opponent_round_avg) else 0

                if team_round_avg > opponent_round_avg:
                    if round['outcome'] == "Win":
                        high_point_wins += 1
                    elif round['outcome'] == "Loss":
                        high_point_losses += 1
                elif team_round_avg < opponent_round_avg:
                    if round['outcome'] == "Win":
                        low_point_wins += 1
                    elif round['outcome'] == "Loss":
                        low_point_losses += 1

                avg_team_speaks.append(team_round_avg)

                # Add to tournament speaker records
                result_id = round['result']['id']
                if result_id not in tourn_to_speaks:
                    tourn_to_speaks[result_id] = []
                tourn_to_speaks[result_id].append(team_round_avg)

                # Record relative stats
                if round['opponentId'] in team_2_speaks and len(team_2_speaks[round['opponentId']]) > 1:
                    opp_speaker_impact.append(
                        opponent_round_avg - statistics.mean(team_2_speaks[round['opponentId']]))
                if 'records' not in round:
                    continue
                for record in round['records']:
                    if record['judgeId'] not in judge_2_speaks or record['judgeId'] not in judge_2_team_avg_points or not len(judge_2_speaks[record['judgeId']]):
                        continue
                    speaks_above_judging.append(
                        judge_2_team_avg_points[record['judgeId']] - statistics.mean(judge_2_speaks[record['judgeId']]))
                speaks_above_average.append(team_round_avg - avg_speaks)

        exp_wins = 0

        for wp in x_wp:
            if wp > 0.5:
                exp_wins += 1

        if not len(rounds) or not len(x_wp):
            continue
        wins_above_replacement = wins_with_exp_wp_recorded - exp_wins
        win_pct_above_replacement = (
            wins_with_exp_wp_recorded/len(rounds)) - (exp_wins/len(rounds))
        x_wp = statistics.mean(x_wp)

        if len(speaks_above_average) and len(speaks_above_judging):
            speaks_above_average = statistics.mean(speaks_above_average) if len(speaks_above_average) else None
            speaks_above_judging = statistics.mean(speaks_above_judging) if len(speaks_above_judging) else None
            opp_speaker_impact = statistics.mean(opp_speaker_impact) if len(opp_speaker_impact) else None
        else:
            speaks_above_average = None
            speaks_above_judging = None
            opp_speaker_impact = None
            high_point_wins = None
            high_point_losses = None
            low_point_wins = None
            low_point_losses = None

        try:
            true_wp = (prelim_wins + prelim_losses)/(prelim_wins + prelim_losses +
                                                     elim_wins + elim_losses) + elim_wins/(elim_wins + elim_losses) * 0.1
        except Exception:
            true_wp = (prelim_wins + prelim_losses) / \
                (prelim_wins + prelim_losses + elim_wins + elim_losses)

        true_wp = 1 if true_wp > 1 else true_wp

        if len(competitor_2_speaks.keys()) >= 2:
            intra_team_speak_std = []
            for _, speaks in competitor_2_speaks.items():
                intra_team_speak_std.append(statistics.mean(speaks))
            intra_team_speak_std = sorted(intra_team_speak_std, reverse=True)
            d_sp = intra_team_speak_std[0] - intra_team_speak_std[-1]
            intra_team_speak_std = statistics.stdev(intra_team_speak_std)

        for tourn_speaks in tourn_to_speaks.values():
            speaks_1hl.append(_hi_lo_avg(tourn_speaks, 1))
            speaks_2hl.append(_hi_lo_avg(tourn_speaks, 2))

        speaks_1hl = list(filter(lambda x: x != None, speaks_1hl))
        speaks_2hl = list(filter(lambda x: x != None, speaks_2hl))

        speaks_1hl = statistics.mean(speaks_1hl) if len(speaks_1hl) else None
        speaks_2hl = statistics.mean(speaks_2hl) if len(speaks_2hl) else None

        statistics_body = {
            'xwp': x_wp,
            'war': wins_above_replacement,
            'wpar': win_pct_above_replacement,
            'stdSpks': overall_speak_std,
            'stdCSpks': intra_team_speak_std,
            'dSpksP': d_sp,
            'saa': speaks_above_average,
            'saj': speaks_above_judging,
            'hpw': high_point_wins,
            'hpl': high_point_losses,
            'lpw': low_point_wins,
            'lpl': low_point_losses,
            'osi': opp_speaker_impact,
            'udw': underdog_wins,
            'udl': underdog_losses,
            'fvw': favorite_wins,
            'fvl': favorite_losses,
            'tourns': tourns,
            'rounds': num_rounds,
            'prelimWins': prelim_wins,
            'prelimLosses': prelim_losses,
            'prelimSplits': prelim_splits,
            'elimWins': elim_wins,
            'elimLosses': elim_losses,
            'flowWins': flow_record[0],
            'flowLosses': flow_record[1],
            'layWins': lay_record[0],
            'layLosses': lay_record[1],
            'progWins': prog_record[0],
            'progLosses': prog_record[1],
            'tradWins': trad_record[0],
            'tradLosses': trad_record[1],
            'totalSpks': sum(avg_team_speaks),
            'avgSpks': sum(avg_team_speaks)/len(avg_team_speaks) if len(avg_team_speaks) else None,
            # 'bids': team_2_bids[result['teamId']],
            'spks1HL': speaks_1hl,
            'spks2HL': speaks_2hl,
            'pwp': prelim_wins / (prelim_wins + prelim_losses) if prelim_wins + prelim_losses != 0 else 0,
            'twp': ((prelim_wins / (prelim_wins + prelim_losses)) if prelim_wins + prelim_losses != 0 else 0) + ((0.1 * elim_wins / (elim_wins + elim_losses)) if elim_wins + elim_losses != 0 else 0)
        }

        create_body = statistics_body
        create_body['team'] = {
            'connect': {
                'id': teamId
            }
        }
        create_body['circuit'] = {
            'connect': {
                'id': circuit
            }
        }
        create_body['season'] = {
            'connect': {
                'id': season
            }
        }
        if teamId not in team_2_otr:
            continue

        create_body['otr'] = team_2_otr[teamId]

        res = requests.post(f'{API_BASE}/rankings/teams/advanced/upsert', json={
            'where': {
                'teamId_circuitId_seasonId': {
                    'teamId': teamId,
                    'circuitId': circuit,
                    'seasonId': season
                }
            },
            'update': statistics_body,
            'create': create_body,
        })

        if res.status_code != 200:
            print({
                    'teamId': teamId,
                    'circuitId': circuit,
                    'seasonId': season
                })

def update_all_stats(job_id: int | None = None):
    circuits = requests.get(f"{API_BASE}/circuits?expand=seasons").json()

    for circuit in circuits:
        for season in circuit['seasons']:
            lprint(job_id, "Info", message=f"Updating (circuit, season) = ({circuit['id']}, {season['id']})")
            try:
                _update_scoped_stats(job_id, season['id'], circuit['id'])
                time.sleep(0.5)
            except Exception:
                lprint(job_id, "Error", message=f"Failed updating (circuit, season) = ({circuit['id']}, {season['id']})")

def update_stats(job_id: int | None, tab_event_id: int):
    """_summary_

    Args:
        tab_event_id (int): _description_
    """

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

    for circuit in circuits:
        lprint(job_id, "Info", message=f"Updating stats for season {season} and circuit {circuit}")
        _update_scoped_stats(job_id, season, circuit)


if __name__ == "__main__":
    update_stats(242828)
