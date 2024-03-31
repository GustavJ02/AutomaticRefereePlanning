import pandas as pd
from datetime import datetime
import json
import ast
from docplex.mp.model import Model
import tracemalloc
import os


class NoSoulutionFound(Exception):
    ...
    pass


class ColleagueError(Exception):
    ...
    pass


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
    for i, date in enumerate(dataframe['Date']):
        ids.append(f"{date.day}-{i}")

    dataframe['id'] = ids


def setRefProperties(dataframe, groups):
    nrs = []
    level = []

    for group in dataframe['Group']:
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
    for time1, place1, id1 in zip(dataframe['Time'], dataframe['Field'], dataframe['id']):
        for time2, place2, id2 in zip(dataframe['Time'], dataframe['Field'], dataframe['id']):
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

        result.at[index, 'Referee 1'] = ref1
        if ref2 != '':
            result.at[index, 'Referee 2'] = ref2

    result = result.drop(['id', 'nrOfRefs', 'reqLevel'], axis=1, inplace=True)


def countAvalibleRefs(referee_dict, i):
    sum = 0
    for referee in referee_dict.keys():
        sum += referee_dict[referee]['Available'][i]

    return sum


def calculateAvrage(referee_dict, gamesInDay):
    result = []

    for i, day in enumerate(gamesInDay):
        avalibleRefs = countAvalibleRefs(referee_dict, i)
        average = 2 * len(day) / avalibleRefs
        result.append(round(average))

    return result


def optimize(referee_dict, games_dict, unAllowedPairs, result, colleagues, gamesInDay, finalGames, gamesOnDayAndField):

    intendedAvrage = calculateAvrage(referee_dict, gamesInDay)

    # ----------------------------------------------------------------
    # Model object
    # ----------------------------------------------------------------
    mod = Model(name='referee_assignment', log_output=True)
    mod.parameters.mip.tolerances.mipgap = 0.29

    # ----------------------------------------------------------------
    # Model data
    # ----------------------------------------------------------------
    print('Reading data')
    print(f'Referees: {referee_dict}', end='\n\n')

    print(f'Colleagues: {colleagues}', end='\n\n')

    referees = list(referee_dict.keys())
    games = list(games_dict.keys())
    daysRange = range(len(gamesInDay))

    gamesNotPool = [x for x in games if games_dict[x]['reqLevel'] != 1]

    notFinalGamesOnDayAndField = []

    for t in daysRange:
        tmpList1 = []
        for field in range(len(gamesOnDayAndField[t])):
            tmpList = []
            for i in range(len(gamesOnDayAndField[t][field])):
                if gamesOnDayAndField[t][field][i] not in finalGames:
                    tmpList.append(gamesOnDayAndField[t][field][i])

            tmpList1.append(tmpList)

        notFinalGamesOnDayAndField.append(tmpList1)

    M = 10000

    # ----------------------------------------------------------------
    # Variables
    # ----------------------------------------------------------------
    print('Creating variables')
    officiates = mod.binary_var_matrix(referees, games, name='officiates')

    # ----------------------------------------------------------------
    # Auxilrary Variables
    # ----------------------------------------------------------------
    above_avg = mod.continuous_var_matrix(referees, daysRange, lb=0, name='above_avg')
    below_avg = mod.continuous_var_matrix(referees, daysRange, lb=0, name='below_avg')
    officiatesLastAndFirst = mod.binary_var_matrix(referees, daysRange, name='officiatesLastAndFirst')

    # ----------------------------------------------------------------
    # Objective
    # ----------------------------------------------------------------
    print('Buiding objective')
    mod.minimize(sum(above_avg[ref, t] + below_avg[ref, t] + 10 * officiatesLastAndFirst[ref, t] for ref in referees for t in daysRange))

    # ----------------------------------------------------------------
    # Constraints
    # ----------------------------------------------------------------
    print('Building constraints')

    # Ensure that each game has the correct number of referees
    for g in games:
        mod.add_constraint(sum(officiates[ref, g] for ref in referees) == games_dict[g]["nrOfRefs"])

    for g in games:
        for ref in referees:
            mod.add_constraint(referee_dict[ref]["Level"] - games_dict[g]["reqLevel"] * officiates[ref, g] >= 0)  # Ensure that each referee has the correct level for each level
            mod.add_constraint(officiates[ref, g] * (referee_dict[ref]["Level"] - games_dict[g]["reqLevel"]) <= 3)  # Ensure that each referee is not overqualified for the game they officiates

    # Ensure that referees are not assigned two simultaions games and allow ample time between games on different fields
    for ref in referees:
        for g1, g2 in unAllowedPairs:
            mod.add_constraint(officiates[ref, g1] + officiates[ref, g2] <= 1)

    # Ensure that referees dont officiate more than 4 consecutive games
    for t in daysRange:
        for index in range(len(gamesInDay[t]) - 4):
            for ref in referees:
                mod.add_constraint(sum(officiates[ref, game] for game in gamesInDay[t][index:index + 5]) <= 4)

    # Ensure that referees officiates a game that is not Poolspel with their collegue if they have one
    for g in gamesNotPool:
        for ref1, ref2 in colleagues:
            mod.add_constraint(officiates[ref1, g] - officiates[ref2, g] == 0)

    # Calculate deviation from avrage for objective and dont assign referees games if they are not avalible
    for ref in referees:
        for t in daysRange:
            mod.add_constraint(sum(officiates[ref, g] for g in gamesInDay[t]) - intendedAvrage[t] == above_avg[ref, t] - below_avg[ref, t])

            mod.add_constraint(sum(officiates[ref, g] for g in gamesInDay[t]) <= M * referee_dict[ref]['Available'][t])

    # Ensure that referees are only assigned one final
    for ref in referees:
        mod.add_constraint(sum(officiates[ref, game] for game in finalGames) <= 1)  # Max one final per referee

    # Add penalty to goal function if referees officiate both first and last game in one day.
    for t in daysRange:
        for field in range(len(gamesOnDayAndField[t])):
            for otherField in range(len(gamesOnDayAndField[t])):
                for ref in referees:
                    mod.add_constraint(officiates[ref, gamesOnDayAndField[t][field][0]] + officiates[ref, gamesOnDayAndField[t][otherField][-1]] <= 1 + officiatesLastAndFirst[ref, t])

    # Makes sure the referees officiates at least two consecutive games
    for t in daysRange:
        for field in range(len(notFinalGamesOnDayAndField[t])):
            for ref in referees:
                mod.add_constraint(officiates[ref, notFinalGamesOnDayAndField[t][field][0]] == officiates[ref, notFinalGamesOnDayAndField[t][field][1]])
                mod.add_constraint(officiates[ref, notFinalGamesOnDayAndField[t][field][-2]] == officiates[ref, notFinalGamesOnDayAndField[t][field][-1]])
                for index in range(1, len(notFinalGamesOnDayAndField[t][field]) - 1):
                    mod.add_constraint(sum(officiates[ref, game] for game in notFinalGamesOnDayAndField[t][field][index-1:index + 2])
                                       >= 2 * officiates[ref, notFinalGamesOnDayAndField[t][field][index]])

    # ----------------------------------------------------------------
    # Solve the model
    # ----------------------------------------------------------------
    print(f'Number of constraints: {mod.number_of_constraints}')

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

        for ref in referees:
            for t in daysRange:
                above_avg_value = sol.get_value(above_avg[ref, t])
                below_avg_value = sol.get_value(below_avg[ref, t])

                if above_avg_value > 0:
                    if above_avg_value <= 1:
                        color_code = '\033[92m'  # Green
                    elif above_avg_value <= 3:
                        color_code = '\033[93m'  # Yellow
                    else:
                        color_code = '\033[91m'  # Red

                    print(f'{color_code}{ref} is above average in day {t} by {above_avg_value}\033[0m')

                elif below_avg_value > 0:
                    if below_avg_value <= 1:
                        color_code = '\033[92m'  # Green
                    elif below_avg_value <= 3:
                        color_code = '\033[93m'  # Yellow
                    else:
                        color_code = '\033[91m'  # Red

                    print(f'{color_code}{ref} is below average in day {t} by {below_avg_value}\033[0m')

                if sol.get_value(officiatesLastAndFirst[ref, t]) > 0.5:
                    print(f'\033[91m{ref} is assigned both first and last game in day {t}\033[0m')

        print(f"Objective value: {mod.objective_value}")

        combine(refs_sol, result)
    else:
        print("No solution found")
        raise NoSoulutionFound


def extractColleagues(referees):
    pairs = []
    for ref1, ref2 in zip(referees['Referee'], referees['Colleague']):
        if not pd.isna(ref2):
            pairs.append([ref1, ref2])

    referee_set = set()  # To store referees encountered
    duplicates = set()   # To store referees appearing in multiple pairs

    noDuplication = list(set(tuple(sorted(pair)) for pair in pairs))

    for pair in noDuplication:
        for referee in pair:
            if referee in referee_set:
                duplicates.add(referee)
            else:
                referee_set.add(referee)

    if duplicates:
        print(f"\033[91mThe following referees appear in multiple pairs: {', '.join(duplicates)}\033[0m")
        raise ColleagueError
    else:
        print("\033[92mNo referee appears in multiple pairs.\033[0m")

    return noDuplication


def populateGamesInDay(games):
    result = []
    for dataframe in games:
        result.append(dataframe['id'].tolist())

    return result


def findFinals(games):
    finals = []
    for index, row in games.iterrows():
        if row['Round'] == 'Final':
            finals.append(row['id'])

    return finals


def findGamesOnDayAndField(games):
    result = []

    for day in games:
        groupedDf = day.groupby('Field')['id'].apply(list).reset_index()
        thisList = groupedDf['id'].tolist()
        result.append(thisList)

    return result


def main():
    games, referees, groups = read_data()

    colleagues = extractColleagues(referees)

    days = len(games)

    [generateID(day) for day in games]

    [setRefProperties(day, groups) for day in games]

    unAllowedPairs = []

    [generateUnAllowedPairs(day, unAllowedPairs) for day in games]

    gamesOnDayAndField = findGamesOnDayAndField(games)

    unAllowedPairs = list(set(tuple(sorted(pair)) for pair in unAllowedPairs))

    gamesInDay = populateGamesInDay(games)

    games = pd.concat(games, ignore_index=True)

    finalGames = findFinals(games)

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
            'Available': available_value
        }
    json_data = {
        'days': days,
        'games': games_dict,
        'referees': referee_dict,
        'colleagues': colleagues,
        'gamesInDay': gamesInDay,
        'finalGames': finalGames,
        'unAllowedPairs': unAllowedPairs
    }

    try:
        os.makedirs('data')
    except FileExistsError:
        pass

    with open('data\data.json', 'w') as json_file:
        json.dump(json_data, json_file, indent=4)

    print(f'Number of unallowed pairs: {len(unAllowedPairs)}')
    try:
        optimize(referee_dict, games_dict, unAllowedPairs, games, colleagues, gamesInDay, finalGames, gamesOnDayAndField)

        print(games)

        games['Date'] = pd.to_datetime(games['Date']).dt.date

        uniqeDates = games['Date'].unique()

        with pd.ExcelWriter('result.xlsx', engine='xlsxwriter') as writer:
            for date in uniqeDates:
                filtered_df = games[games['Date'] == date]
                filtered_df.to_excel(writer, sheet_name=str(date), index=False)

    except NoSoulutionFound:
        print(f'\033[91mUnable to find a soulution.\033[0m')


if __name__ == "__main__":
    tracemalloc.start()

    try:
        main()
    except ColleagueError:
        print("\033[91mUnsolvable error in colleague check. Stopping program. Please confirm the colleages in the input data.\033[0m")

    print(f'Memory usage (current, peak): {tracemalloc.get_traced_memory()}')

    tracemalloc.stop()
