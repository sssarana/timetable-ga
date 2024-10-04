import pandas as pd
import random
from openpyxl.utils import get_column_letter

# Load Excel data using pandas
groups_df = None
while groups_df is None:
    file_path = input("Enter the name of the file: ")
    try:
        groups_df = pd.read_excel(file_path, sheet_name='Groups')
        modules_df = pd.read_excel(file_path, sheet_name='Modules')
        rooms_df = pd.read_excel(file_path, sheet_name='Rooms')
    except FileNotFoundError:
        print(f"File {file_path} doesn't exist here, try again")

# Extract groups, modules, and rooms data
GROUPS = groups_df['GroupName'].tolist()

GROUP_MODULES = {}
for _, row in modules_df.iterrows():
    group = row['Group']
    module_key = row['Module']
    module_name = row['ModuleName']
    lecturers = [lecturer.strip() for lecturer in row['Lecturers'].split(',')]
    if group not in GROUP_MODULES:
        GROUP_MODULES[group] = {}
    GROUP_MODULES[group][module_key] = {"name": module_name, "lecturers": lecturers}

LECTURE_ROOMS = rooms_df['LectureRooms'].dropna().tolist()
LAB_ROOMS = rooms_df['LabRooms'].dropna().tolist()

# Define available time slots
DAYS = ["Mon", "Tue", "Wed", "Thu", "Fri"]
HOURS = [(9, 10), (10, 11), (11, 12), (12, 1), (1, 2), (2, 3), (3, 4), (4, 5), (5, 6)]
TIME_SLOTS = [f"{day}_{start}-{end}" for day in DAYS for (start, end) in HOURS]
TIME_SLOTS_PRIORITIZED = [f"{day}_{start}-{end}" for day in DAYS[:-1] for (start, end) in HOURS]  # Monday to Thursday
FRIDAY_SLOTS = [f"Fri_{start}-{end}" for (start, end) in HOURS]

# Configurations
LECTURES_PER_MODULE = 2
LABS_PER_MODULE = 1

# Variables: Each lecture and lab session is a variable
VARIABLES = []
for group, modules in GROUP_MODULES.items():
    for module_key, module_info in modules.items():
        for lecture in range(1, LECTURES_PER_MODULE + 1):
            VARIABLES.append({
                "group": group,
                "module": module_key,
                "session": f"Lecture{lecture}",
                "module_name": module_info["name"],
                "lecturers": module_info["lecturers"]
            })
        VARIABLES.append({
            "group": group,
            "module": module_key,
            "session": "Lab",
            "module_name": module_info["name"],
            "lecturers": module_info["lecturers"]
        })

DOMAINS = {}
for var in VARIABLES:
    var_key = f"{var['group']}_{var['module']}_{var['session']}"
    if var["session"] == "Lab":
        # Labs must take two consecutive time slots, preferably Monday to Thursday
        lab_times = [(TIME_SLOTS_PRIORITIZED[i], TIME_SLOTS_PRIORITIZED[i + 1]) for i in range(len(TIME_SLOTS_PRIORITIZED) - 1)
                     if TIME_SLOTS_PRIORITIZED[i].split("_")[0] == TIME_SLOTS_PRIORITIZED[i + 1].split("_")[0]]
        friday_lab_times = [(FRIDAY_SLOTS[i], FRIDAY_SLOTS[i + 1]) for i in range(len(FRIDAY_SLOTS) - 1)]

        DOMAINS[var_key] = [(time1, time2, room, lecturer) for (time1, time2) in lab_times for room in LAB_ROOMS for lecturer in var["lecturers"]] + \
                           [(time1, time2, room, lecturer) for (time1, time2) in friday_lab_times for room in LAB_ROOMS for lecturer in var["lecturers"]]
    else:
        # Lectures take one time slot, preferably Monday to Thursday
        DOMAINS[var_key] = [(time, room, lecturer) for time in TIME_SLOTS_PRIORITIZED for room in LECTURE_ROOMS for lecturer in var["lecturers"]] + \
                           [(time, room, lecturer) for time in FRIDAY_SLOTS for room in LECTURE_ROOMS for lecturer in var["lecturers"]]

# Generate variable key
def get_var_key(var):
    return f"{var['group']}_{var['module']}_{var['session']}"

# Fitness function to evaluate the quality of an individual schedule
def fitness(individual):
    penalty = 0
    scheduled_sessions = {}
    for var_key, value in individual.items():
        if "Lab" in var_key:
            time1, time2, room, lecturer = value

            # Deal with group conflicts
            group = var_key.split('_')[0]
            if (group, time1) in scheduled_sessions or (group, time2) in scheduled_sessions:
                penalty += 1
            else:
                scheduled_sessions[(group, time1)] = True
                scheduled_sessions[(group, time2)] = True

            # Deal with room conflicts
            if (room, time1) in scheduled_sessions or (room, time2) in scheduled_sessions:
                penalty += 1
            else:
                scheduled_sessions[(room, time1)] = True
                scheduled_sessions[(room, time2)] = True

            # Deal with lecturer conflicts
            if (lecturer, time1) in scheduled_sessions or (lecturer, time2) in scheduled_sessions:
                penalty += 1
            else:
                scheduled_sessions[(lecturer, time1)] = True
                scheduled_sessions[(lecturer, time2)] = True

            # Penalise use of Friday
            if "Fri" in time1 or "Fri" in time2:
                penalty += 2

        else:
            time, room, lecturer = value

            # Deal with group conflicts
            group = var_key.split('_')[0]
            if (group, time) in scheduled_sessions:
                penalty += 1
            else:
                scheduled_sessions[(group, time)] = True

            # Deal with room conflicts
            if (room, time) in scheduled_sessions:
                penalty += 1
            else:
                scheduled_sessions[(room, time)] = True

            # Deal with lecturer conflicts
            if (lecturer, time) in scheduled_sessions:
                penalty += 1
            else:
                scheduled_sessions[(lecturer, time)] = True

            # Penalise use of Friday
            if "Fri" in time:
                penalty += 2

    return -penalty

# Initialise a population of schedules
def initialize_population(size):
    population = []
    for _ in range(size):
        individual = {}
        for var in VARIABLES:
            var_key = get_var_key(var)
            individual[var_key] = random.choice(DOMAINS[var_key])
        population.append(individual)
    return population

# Crossover function for genetic algorithm
def crossover(parent1, parent2):
    crossover_point = len(parent1) // 2
    child1 = {**dict(list(parent1.items())[:crossover_point]), **dict(list(parent2.items())[crossover_point:])}
    child2 = {**dict(list(parent2.items())[:crossover_point]), **dict(list(parent1.items())[crossover_point:])}
    return child1, child2

# Mutation function for genetic algorithm
def mutate(individual, mutation_rate=0.01):
    for var_key in individual:
        if random.random() < mutation_rate:
            individual[var_key] = random.choice(DOMAINS[var_key])
    return individual

def genetic_algorithm(pop_size, generations):
    population = initialize_population(pop_size)
    for generation in range(generations):
        population = sorted(population, key=lambda ind: fitness(ind), reverse=True)
        new_population = population[:pop_size // 2]

        # Create offspring
        for i in range(0, len(new_population), 2):
            parent1, parent2 = new_population[i], new_population[min(i + 1, len(new_population) - 1)]
            child1, child2 = crossover(parent1, parent2)
            new_population.extend([mutate(child1), mutate(child2)])

        population = new_population[:pop_size]
    return max(population, key=lambda ind: fitness(ind))

# Find an optimal timetable
solution = genetic_algorithm(pop_size=50, generations=100)

if solution:
    # Empty dictionary to store the timetables for each group
    timetable = {group: pd.DataFrame(index=DAYS, columns=[f"{start}-{end}" for (start, end) in HOURS]) for group in GROUPS}

    # Put the solution into the timetable dictionary
    for var_key, value in solution.items():
        group, module, session = var_key.split("_")
        if "Lab" in session:
            time1, time2, room, lecturer = value
            times = [time1, time2]
        else:
            time, room, lecturer = value
            times = [time]

        for time in times:
            day, hour = time.split("_")
            session_info = f"{module} ({session}) in {room} by {lecturer}"
            if pd.isna(timetable[group].loc[day, hour]):
                timetable[group].loc[day, hour] = session_info
            else:
                timetable[group].loc[day, hour] += f"; {session_info}"

    # Export each group's timetable as a separate sheet in the Excel file
    with pd.ExcelWriter('timetable_output_matrix.xlsx', engine='openpyxl') as writer:
        for group, df in timetable.items():
            df.to_excel(writer, sheet_name=f"{group}")

            # Open the workbook to modify column widths (looks better)
            workbook = writer.book
            worksheet = workbook[f"{group}"]

            for col_num, col in enumerate(df.columns, 2):
                max_length = max(
                    df[col].astype(str).map(len).max(),  # Length of the longest entry in this column
                    len(str(col))  # Length of the column header
                ) + 2  # padding
                column_letter = get_column_letter(col_num)
                worksheet.column_dimensions[column_letter].width = max_length

            # Set row heights based on the longest content in the index column
            for row_num, day in enumerate(df.index, 2):  # Start at 2 for Excel indexing
                max_length = max(len(str(day)), df.loc[day].astype(str).map(len).max()) + 2
                worksheet.row_dimensions[row_num].height = max_length

    print("Timetable has been successfully exported to 'timetable_output_matrix.xlsx'")

else:
    print("No solution possible.")
