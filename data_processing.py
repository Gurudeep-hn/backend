""" Defined function for static initial data for Production Planning """

#importing libraries
import pandas as pd
import openpyxl as xl
from datetime import datetime, timedelta
import numpy as np




def data_excel_json(excel_sheet):
    """ convert excel into json """
    data_excel = xl.load_workbook(excel_sheet)
    data = {}
    sheet_name = data_excel.sheetnames
    for sheet in sheet_name:
        wb_sheet = data_excel[sheet]
        cell_values = wb_sheet.values
        df =  pd.DataFrame(cell_values, columns=next(cell_values))
        df.iloc[:, 0] = df.iloc[:, 0].apply(lambda x : x.strip())
        df.index = df.iloc[:, 0]
        df.drop(columns = df.columns[0], inplace=True)
        data[sheet] = df.T.to_dict()
    return data

def json_to_df(json_data):
    """ convert json into excel """
    dict_data = {}
    for key in json_data.keys():
        dict_data[key] = pd.DataFrame(json_data.get(key)).T
    return dict_data


def process_final_schedule(final_schedule):
    # Assuming final_schedule is a list of dictionaries representing tasks
    job_task_counts = {}
    
    for task in final_schedule['data']:
        job = task['Resource']
        if job in job_task_counts:
            job_task_counts[job] += 1
        else:
            job_task_counts[job] = 1
        task['TaskId'] = str(job_task_counts[job])
    
    return final_schedule



def process_final_schedule_updated(final_schedule):
    updated_schedule = []
    task_id_counter = 1  # Starting ID for tasks
    
    for task in final_schedule['data']:
        # Extract and convert start and finish times
        start_time = datetime.strptime(task['Start'], '%Y-%m-%dT%H:%M:%S')
        end_time = datetime.strptime(task['Finish'], '%Y-%m-%dT%H:%M:%S')

        # Extract job number and task number from Resource and TaskId
        job_number = task['Resource'].split(' ')[-1]  # Assuming the job number is the last part of the Resource string
        task_number = task['TaskId']

        # Construct the new task format
        new_task = {
            'Id': task_id_counter,
            'Subject': f"Task {task_number.zfill(2)} of Job {job_number}",  # Constructing Subject in "Task 02 of Job 01" format
            'Description': task['Task'],
            'StartTime': int(start_time.timestamp() * 1000),  # Convert to milliseconds for JavaScript
            'EndTime': int(end_time.timestamp() * 1000),  # Convert to milliseconds for JavaScript
            'RoomId': int(task_number)  # Assuming TaskId is suitable for RoomId
        }

        updated_schedule.append(new_task)
        task_id_counter += 1  # Incrementing the task ID counter for the next task

    return updated_schedule

def parse_duration_to_timedelta(duration_str):
    """Parse a duration string into a timedelta object."""
    days, time_part = 0, duration_str
    if 'day' in duration_str:
        days_part, time_part = duration_str.split(',', 1)
        days = int(days_part.split()[0])
    hours, minutes, seconds = map(int, time_part.split(':'))
    return timedelta(days=days, hours=hours, minutes=minutes, seconds=seconds)

def get_job_machine_details(sequence_best, process_time, machine_sequence, num_jobs, num_machines, start_datetime_str='2024-02-01 08:00:00'):
    start_datetime = datetime.strptime(start_datetime_str, '%Y-%m-%d %H:%M:%S')
    
    job_counts = {job: 0 for job in range(num_jobs)}
    machine_end_times = {machine: 0 for machine in range(1, num_machines + 1)}

    job_details = []

    for job_index in sequence_best:
        task_index = job_counts[job_index]
        machine_index = machine_sequence[job_index][task_index] - 1
        duration = process_time[job_index][task_index]

        start_time_minutes = max(machine_end_times[machine_index + 1], sum(process_time[job_index][:task_index]))
        end_time_minutes = start_time_minutes + duration

        actual_start_time = start_datetime + timedelta(minutes=start_time_minutes)
        actual_end_time = start_datetime + timedelta(minutes=end_time_minutes)

        machine_end_times[machine_index + 1] = end_time_minutes
        job_counts[job_index] += 1

        job_details.append({
            'Job': job_index + 1,
            'Subject': f"Task {task_index + 1} of Job {job_index + 1}",
            'Description': f"Machine {machine_index + 1}",
            'Id': machine_index + 1,
            'StartTime': int(actual_start_time.timestamp() * 1000),
            'EndTime': int(actual_end_time.timestamp() * 1000),
            'RoomId': machine_index + 1
        })

    return job_details



def convert_numpy(obj):
    if isinstance(obj, np.integer):
        return int(obj)
    elif isinstance(obj, np.floating):
        return float(obj)
    elif isinstance(obj, np.ndarray):
        return obj.tolist()
    elif isinstance(obj, dict):
        return {key: convert_numpy(value) for key, value in obj.items()}
    elif isinstance(obj, list):
        return [convert_numpy(item) for item in obj]
    else:
        return obj