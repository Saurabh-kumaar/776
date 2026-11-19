import openpyxl

# Open the Excel file
try:
    workbook = openpyxl.load_workbook("../12603564.xlsx", data_only=True)
    sheet = workbook["Daily Log"]

except FileNotFoundError:
    print("Excel file not found.")
    exit()

except Exception as error:
    print("Error while opening Excel file:", error)
    exit()

#=================================
# Lists to store the daily data:-
sleep = []
fitness = []
study = []
coding = []
class_time = []
class_attended = []
other_activities = []
total_tracked = []
free_uncounted = []
day_feeling = []
satisfaction_level = []
energy_level = []
academic_activity = []
experience = []

# Lists for relationship analysis
coding_energy = []
sleep_energy = []
study_satisfaction = []

# Store the dates
dates = []

#=================================
# Find the average of the values
def calculate_average(data):
    total = 0
    count = 0

    for value in data:
        total = total + value
        count = count + 1

    if count == 0:
        return 0

    average = total / count

    return average

#===================================
# Calculate Tech Productivity Index
def calculate_tpi(coding):
    return calculate_average(coding)

#====================================
# Calculate Academic Activity Index
def calculate_aai(academic_activity):
    return calculate_average(academic_activity)

#====================================
# Calculate Physical Activity Index
def calculate_phai(fitness):
    return calculate_average(fitness)

# ===================================
# Calculate Sleep and Recovery Index
def calculate_sri(sleep):
    return calculate_average(sleep)

# ===================================
# Time Utilization Index :-
def calculate_tui(total_tracked):
    return calculate_average(total_tracked)


#====================================
# Calculate Activity Balance Index:-
def calculate_abi(free_uncounted):
    return calculate_average(free_uncounted)


#====================================
# Calculate Experience Index:-
def calculate_ei(experience):
    return calculate_average(experience)

#=====================================
# Correction two values:-
def calculate_correlation(data):
    x_total = 0
    y_total = 0
    count = 0

    for x, y in data:
        x_total = x_total + x
        y_total = y_total + y
        count = count + 1

    if count == 0:
        return 0

    x_average = x_total / count
    y_average = y_total / count

    numerator = 0
    x_difference = 0
    y_difference = 0

    for x, y in data:
        x_diff = x - x_average
        y_diff = y - y_average

        numerator = numerator + (x_diff * y_diff)
        x_difference = x_difference + (x_diff * x_diff)
        y_difference = y_difference + (y_diff * y_diff)

    denominator = (x_difference * y_difference) ** 0.5

    if denominator == 0:
        return 0

    correlation = numerator / denominator

    return correlation

#============================================
# Calculate Data Continuity Index
def calculate_dci(valid_days, expected_days):
    return (valid_days / expected_days) * 100

#=============================================
# Calculate the final Personal Activity Index
def calculate_pai(tpi, aai, phai, sri, tui, ei, dci):
    return (0.15 * tpi) + (0.20 * aai) + (0.15 * phai) + (0.20 * sri) + (0.15 * tui) + (0.10 * ei) + (0.05 * dci)

#=============================================
# Read data from the Excel rows
for row in range(7, sheet.max_row + 1):

    # Read the date
    date_value = sheet.cell(row, 1).value

    if date_value is None:
        continue

    dates.append(date_value)

    # Read sleep time
    value = sheet.cell(row, 2).value
    if value is not None:
        sleep.append(value)

    # Read fitness time
    fitness_value = sheet.cell(row, 3).value
    if fitness_value is not None:
        fitness.append(fitness_value)

    # Read study time
    study_value = sheet.cell(row, 4).value
    if study_value is not None:
        study.append(study_value)

    # Read coding time
    coding_value = sheet.cell(row, 5).value
    if coding_value is not None:
        coding.append(coding_value)

    # Read class time
    class_value = sheet.cell(row, 6).value
    if class_value is not None:
        class_time.append(class_value)

    # Add study and class time for academic activity
    if study_value is not None and class_value is not None:
        academic_activity.append(study_value + class_value)

    # Read classes attended
    class_attended_value = sheet.cell(row, 7).value
    if class_attended_value is not None:
        class_attended.append(class_attended_value)

    # Read other activities
    other_activities_value = sheet.cell(row, 8).value
    if other_activities_value is not None:
        other_activities.append(other_activities_value)

    # Read total tracked time
    total_tracked_value = sheet.cell(row, 9).value
    if total_tracked_value is not None:
        total_tracked.append(total_tracked_value)

    # Read free or unaccounted time
    free_uncounted_value = sheet.cell(row, 10).value
    if free_uncounted_value is not None:
        free_uncounted.append(free_uncounted_value)

    # Convert feeling into a number
    feeling_value = sheet.cell(row, 11).value

    if feeling_value is not None:

        if feeling_value == "Excellent":
            feeling_score = 5
        elif feeling_value == "Good":
            feeling_score = 4
        elif feeling_value == "Neutral":
            feeling_score = 3
        elif feeling_value == "Low":
            feeling_score = 2
        else:
            feeling_score = 1

        day_feeling.append(feeling_score)

    # Convert satisfaction into a number
    satisfaction_level_value = sheet.cell(row, 12).value

    if satisfaction_level_value is not None:

        if satisfaction_level_value == "Very Satisfied":
            satisfaction_score = 5
        elif satisfaction_level_value == "Satisfied":
            satisfaction_score = 4
        elif satisfaction_level_value == "Neutral":
            satisfaction_score = 3
        elif satisfaction_level_value == "Unsatisfied":
            satisfaction_score = 2
        else:
            satisfaction_score = 1

        satisfaction_level.append(satisfaction_score)

    # Convert energy level into a number
    energy_level_value = sheet.cell(row, 13).value

    if energy_level_value is not None:

        if energy_level_value == "High":
            energy_score = 3
        elif energy_level_value == "Medium":
            energy_score = 2
        else:
            energy_score = 1

        energy_level.append(energy_score)

    #=====================================================
    # Calculate experience Index :-

    if satisfaction_level_value is not None and feeling_value is not None and energy_level_value is not None:
        experience.append((feeling_score + satisfaction_score + energy_score) / 3)

    #=====================================================
    # Store coding and energy together:-
    if coding_value is not None and energy_level_value is not None:
        coding_energy.append((coding_value, energy_score))

    #=====================================================
    # Store sleep and energy together:-
    if value is not None and energy_level_value is not None:
        sleep_energy.append((value, energy_score))

    #=====================================================
    # Store study and satisfaction together:-
    if study_value is not None and satisfaction_level_value is not None:
        study_satisfaction.append((study_value, satisfaction_score))


#======================================
# Calculate all activity indexes
tpi = calculate_tpi(coding)
aai = calculate_aai(academic_activity)
phai = calculate_phai(fitness)
sri = calculate_sri(sleep)
tui = calculate_tui(total_tracked)
abi = calculate_abi(free_uncounted)
ei = calculate_ei(experience)
study_average = calculate_average(study)
class_average = calculate_average(class_time)
other_average = calculate_average(other_activities)


#======================================
# Calculate the three relationships
coding_energy_result = calculate_correlation(coding_energy)
sleep_energy_result = calculate_correlation(sleep_energy)
study_satisfaction_result = calculate_correlation(study_satisfaction)

#======================================
# Calculate data continuity
valid_days = len(dates)
expected_days = 36

dci = calculate_dci(valid_days, expected_days)


# Calculate final PAI
pai = calculate_pai(tpi, aai, phai, sri, tui, ei, dci)


#============================================
# Display the activity index results
print("\n----- ACTIVITY INDEX RESULTS -----")  

print(f"avg study/day : {study_average:.2f} min/day")
print(f"avg class/day : {class_average:.2f} min/day")
print(f"avg value other/day : {other_average:.2f} min/day")

#==========================================================

print(f"TPI  (Tech Productivity) : {tpi:.2f} min/day")
print(f"AAI  (Academic Activity)  : {aai:.2f} min/day")
print(f"PhAI (Physical Activity)  : {phai:.2f} min/day")
print(f"SRI  (Sleep & Recovery)   : {sri:.2f} min/day")
print(f"TUI  (Time Utilization)   : {tui:.2f} min/day")
print(f"ABI  (Activity Balance)   : {abi:.2f} min/day")
print(f"EI   (Experience Index)   : {ei:.2f} / 5")
print(f"DCI  (Data Continuity)    : {dci:.2f} %")
print(f"PAI  (Personal Activity)  : {pai:.2f}")


# Display relationship results
print("\n----- RELATIONSHIP ANALYSIS -----")

print(f"Coding - Energy        : {coding_energy_result:.2f}")
print(f"Sleep - Energy         : {sleep_energy_result:.2f}")
print(f"Study - Satisfaction   : {study_satisfaction_result:.2f}") 



#------------- End this project here ---------------------------------
#=====================================================================





