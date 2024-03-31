# Description
This project reads data from Excel and stores that in pandas dataframes. Applies necessary calculations to find lists/dictionaries/dataframes that will be used to formulate a MILP optimization problem and construct constraints for the optimization. The code then uses the IBM CPLEX solver to find the optimal solution to the problem.

# Rules/constraints for planning

- Don't assign referees games if they are not available
- Ensure that each game has the correct number of referees
- Ensure that each referee has the correct qualification for each level
- Ensure that each referee is not overqualified for the game they officiate
- Ensure that referees are not assigned two simoultaions games and allow ample time between games on different fields
- Ample time means more than 1.2 hours between the start of game A on field X and the start of game B on field Y
- Ensure that referees don't officiate more than 4 consecutive games
- Ensure that referees officiate a game with their colleague if they have one
- Ensure that referees are only assigned one final
- Makes sure the referees officiate at least two consecutive games

Also, constraint to formulate an effective objective function
- Calculate deviation from average for objective
- Add penalty to objective if one referee is assigned both the first and last game in one day (effectively first/last two games due to other constraints)
