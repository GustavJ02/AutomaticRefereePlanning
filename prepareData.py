import pandas as pd
from datetime import datetime
import json
import ast
from docplex.mp.model import Model


def read_data():
    file = pd.ExcelFile('InputData.xlsx')

    dataframes = []
    refs = pd.read_excel(file, 'Referees')
    groups = pd.read_excel(file, 'Groups').set_index('Group')['Level'].to_dict()

    refs['Available'] = refs['Available'].apply(ast.literal_eval)

    for sheet in file.sheet_names:
        if 'Day' in sheet:
            dataframes.append(pd.read_excel(file, sheet).dropna(how='all'))

    return dataframes, refs, groups


def generateID(dataframe):
    ids = []
    for i, date in enumerate(dataframe['date']):
        ids.append(f"{date.day}-{i}")

    dataframe['id'] = ids


def setRefProperties(dataframe, groups):
    nrs = []
    level = []

    for group in dataframe['Grupp']:
        if 'Pool' in group:
            nrs.append(1)
        else:
            nrs.append(2)

        for key in groups.keys():
            if key in group:
                level.append(groups[key])

    dataframe['nrOfRefs'] = nrs
    dataframe['reqLevel'] = level


def generateUnAllowedPairs(dataframe, unAllowedPairs):
    for time1, place1, id1 in zip(dataframe['Tid'], dataframe['Spelplan'], dataframe['id']):
        for time2, place2, id2 in zip(dataframe['Tid'], dataframe['Spelplan'], dataframe['id']):
            if place1 != place2:

                time1_dt = datetime.combine(datetime(1900, 1, 1).date(), time1)
                time2_dt = datetime.combine(datetime(1900, 1, 1).date(), time2)

                time_diff = abs((time1_dt - time2_dt).total_seconds() / 3600)
                if time_diff <= 1.2:
                    unAllowedPairs.append([id1, id2])


def combine(refs_sol, result):
    for index, row in result.iterrows():
        key = row['id']

        ref1 = refs_sol[key][0]
        if len(refs_sol[key]) > 1:
            ref2 = refs_sol[key][1]
        else:
            ref2 = ''

        # print(f'Key: {key}, Ref1: {ref1}, Ref2: {ref2}')

        result.at[index, 'Domare 1'] = ref1
        if ref2 != '':
            result.at[index, 'Domare 2'] = ref2

    result = result.drop(['id', 'nrOfRefs', 'reqLevel'], axis=1, inplace=True)


def countAvalibleRefs(referee_dict, i):
    sum = 0
    for referee in referee_dict.keys():
        sum += referee_dict[referee]['available'][i]

    return sum


def calculateAvrage(referee_dict, gamesInDay):
    result = []

    for i, day in enumerate(gamesInDay):
        avalibleRefs = countAvalibleRefs(referee_dict, i)
        average = len(day) / avalibleRefs
        result.append(average)

    return result


def optimize(referee_dict, games_dict, unAllowedPairs, result, colleagues, gamesInDay):

    intendedAvrage = calculateAvrage(referee_dict, gamesInDay)

    # ----------------------------------------------------------------
    # Model object
    # ----------------------------------------------------------------
    mod = Model(name='referee_assignment', log_output=True)

    # ----------------------------------------------------------------
    # Model data
    # ----------------------------------------------------------------
    referees = list(referee_dict.keys())
    games = list(games_dict.keys())
    daysRange = range(len(gamesInDay))

    M = 1000

    # ----------------------------------------------------------------
    # Variables
    # ----------------------------------------------------------------
    officiates = mod.binary_var_matrix(referees, games, name='officiates')

    # ----------------------------------------------------------------
    # Auxilrary Variables
    # ----------------------------------------------------------------
    above_avg = mod.continuous_var_matrix(referees, daysRange, lb=0, name='above_avg')
    below_avg = mod.continuous_var_matrix(referees, daysRange, lb=0, name='below_avg')

    # ----------------------------------------------------------------
    # Objective
    # ----------------------------------------------------------------
    mod.minimize(sum(above_avg[ref, t] + below_avg[ref, t] for ref in referees for t in daysRange))

    # ----------------------------------------------------------------
    # Constraints
    # ----------------------------------------------------------------
    for g in games:
        mod.add_constraint(sum(officiates[ref, g] for ref in referees) == games_dict[g]["nrOfRefs"])

    for g in games:
        for ref in referees:
            mod.add_constraint(referee_dict[ref]["Level"] - games_dict[g]["reqLevel"] * officiates[ref, g] >= 0)  # Ensure that each referee has the correct level for each level
            mod.add_constraint(officiates[ref, g] * (referee_dict[ref]["Level"] - games_dict[g]["reqLevel"]) <= 3)  # Ensure that each referee is not overqualified for the game they officiates

    for ref in referees:
        for g1, g2 in unAllowedPairs:
            mod.add_constraint(officiates[ref, g1] + officiates[ref, g2] <= 1)

    for index, game in enumerate(games[:-4]):
        for ref in referees:
            mod.add_constraint(sum(officiates[ref, games[i]] for i in range(index, index + 4)) <= 4)

    for g in games:
        for ref1, ref2 in colleagues:
            mod.add_constraint(officiates[ref1, g] - officiates[ref2, g] == 0)

    for ref in referees:
        for t in daysRange:
            mod.add_constraint(sum(officiates[ref, g] for g in gamesInDay[t]) - intendedAvrage[t] == above_avg[ref, t] - below_avg[ref, t])

            mod.add_constraint(sum(officiates[ref, g] for g in gamesInDay[t]) <= M * referee_dict[ref]['available'][t])

    # ----------------------------------------------------------------
    # Solve the model
    # ----------------------------------------------------------------
    # print(mod)
    sol = mod.solve()

    # ----------------------------------------------------------------
    # Print the solution
    # ----------------------------------------------------------------
    refs_sol = {}

    if sol:
        print("Solution found:")
        for ref in referees:
            for g in games:
                if sol.get_value(officiates[ref, g]) > 0.5:
                    # print(f"{ref} officiates {g}")
                    if g in refs_sol.keys():
                        refs_sol[g] = [refs_sol[g][0], ref]
                    else:
                        refs_sol[g] = [ref]

        print(f"Objective value: {mod.objective_value}")
        # print(refs_sol)
        combine(refs_sol, result)
    else:
        print("No solution found")


def extractColleagues(referees):
    pairs = []
    for ref1, ref2 in zip(referees['Referee'], referees['colleague']):
        if not pd.isna(ref2):
            pairs.append([ref1, ref2])

    return list(set(tuple(sorted(pair)) for pair in pairs))


def populateGamesInDay(games):
    result = []
    for dataframe in games:
        result.append(dataframe['id'].tolist())

    return result


def main():
    games, referees, groups = read_data()

    colleagues = extractColleagues(referees)

    days = len(games)

    # print(referees)
    # print(colleagues)
    # print(groups)

    [generateID(day) for day in games]

    [setRefProperties(day, groups) for day in games]

    unAllowedPairs = []

    [generateUnAllowedPairs(day, unAllowedPairs) for day in games]

    unAllowedPairs = list(set(tuple(sorted(pair)) for pair in unAllowedPairs))

    # [print(day) for day in games]

    gamesInDay = populateGamesInDay(games)

    games = pd.concat(games, ignore_index=True)
    # print(games)

    games_dict = {}

    for index, row in games.iterrows():
        id_key = row['id']
        nrOfRefs_value = row['nrOfRefs']
        reqLevel_value = row['reqLevel']

        games_dict[id_key] = {
            'nrOfRefs': nrOfRefs_value,
            'reqLevel': reqLevel_value
        }

    referee_dict = {}

    for index, row in referees.iterrows():
        referee_name = row['Referee']
        level_value = row['Level']
        available_value = list(row['Available'])

        # Create a dictionary for each referee name
        referee_dict[referee_name] = {
            'Level': level_value,
            'available': available_value
        }
    json_data = {
        'days': days,
        'games': games_dict,
        'referees': referee_dict,
        'colleagues': colleagues,
        'gamesInDay': gamesInDay,
        'unAllowedPairs': unAllowedPairs
    }

    with open('data\data.json', 'w') as json_file:
        json.dump(json_data, json_file, indent=4)

    print(f'Number of unallowed pairs: {len(unAllowedPairs)}')
    optimize(referee_dict, games_dict, unAllowedPairs, games, colleagues, gamesInDay)

    print(games)
    with pd.ExcelWriter('result.xlsx', engine='xlsxwriter') as writer:
        games.to_excel(writer, sheet_name='games', index=False)


if __name__ == "__main__":
    main()
