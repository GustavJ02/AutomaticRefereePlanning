# Description
This project reads data from excel and stores that in pandas dataframes. Applies neccecary calculations to find lists/dictionaries/dataframes that will be used to formulate a MILP optimization problem and construct constraints for the optimization. The code then uses the IBM CPLEX solver to find the optimal soulution to the problem.

# Rules/constraints for planning

- Ensure that each game has the correct number of referees
- Ensure that each referee has the correct qualification for each level
- Ensure that each referee is not overqualified for the game they officiates
- Ensure that referees are not assigned two simultaions games and allow ample time between games on different fields
- Ample time means more than 1.2 hours between the start of game A on field X and start of game B on field Y
- Ensure that referees dont officiate more than 4 consecutive games
- Ensure that referees officiates a game with their collegue if they have one
- Ensure that referees are only assigned one final
- Makes sure the referees officiates at least two consecutive games

# Also this constraint to formulate an effective objective function
- Calculate deviation from avrage for objective and dont assign referees games if they are not avalible
