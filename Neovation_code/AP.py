import input_file
import worker
import datetime
import numpy as np
import datetime


class bcolors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKCYAN = '\033[96m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'


# Import necessary modules and classes

def months_between(start_date, end_date):
    # Calculate months between two dates
    start_date = datetime.datetime.strptime(start_date, "%d.%m.%Y")
    end_date = datetime.datetime.strptime(end_date, "%d.%m.%Y")
    return (end_date.year - start_date.year) * 12 + end_date.month - start_date.month + 1


class AP:
    def __init__(self):
        # Initialize the Annual Planner with empty lists and variables
        self.aps_number_distributed = []
        self.dates_distributed = []
        self.dates_st = []  # List to store start dates
        self.dates_ft = []  # List to store end dates
        self.intervals = []  # List to store intervals between start and end dates
        self.workers = []  # List to store assigned workers
        self.hours = []  # List to store worked hours
        self.year_start = 0  # Variable to store the start year
        self.year_end = 0  # Variable to store the end year
        self.working_hours = []  # List to store individual working hours
        self.Nr = []  # List to store the number of workers
        self.working_dates_start = []
        self.working_dates_end = []

    def add_dates(self, dates_start_list, dates_end_list):
        # Add start and end dates to the planner
        for dt in dates_end_list:
            self.dates_ft.append(dt)

        for dt in dates_start_list:
            self.dates_st.append(dt)

    def get_hours(self, hours):
        # Set the hours worked for each interval
        self.hours = hours

    def add_Nr(self, Nr):
        # Add the number of workers
        self.Nr = Nr

    def get_smallest_year(self):
        # Determine the smallest year in the provided dates
        smallest_year = 1000000
        for dt in self.dates_st:
            temp_year = int(dt[6:])
            if temp_year < smallest_year:
                smallest_year = temp_year
        self.year_start = smallest_year

    def get_biggest_year(self):
        # Determine the biggest year in the provided dates
        biggest_year = 0
        for dt in self.dates_ft:
            temp_year = int(dt[6:])
            if temp_year > biggest_year:
                biggest_year = temp_year
        self.year_end = biggest_year

    def check_if_same_years(self, ids, repeted_wh_ids, list_begin, list_end):
        # Check if intervals span across multiple years and calculate duration for each year
        index_Nr = 0
        index = 0
        for dt in zip(self.dates_st, self.dates_ft):
            st = dt[0]
            ft = dt[1]
            year_begin = list_begin[index]
            year_end = list_end[index]

            if year_begin == year_end:
                # If start and end dates are in the same year, calculate interval directly
                self.intervals.append([calculate_delta(st, ft)])
                self.working_dates_start.append(st)
                self.working_dates_end.append(ft)

            else:
                # If start and end dates span across multiple years, calculate intervals for each year
                intervals_years = []
                start_year = int(st[6:])
                end_year = int(ft[6:])
                delta_year = end_year - start_year + 1
                unix_st = st
                unix_end = ""
                counter = 0
                for year in range(delta_year):
                    if year > 0:
                        unix_st = "01.01." + str(int(st[6:]) + year)

                    if year == delta_year - 1:
                        unix_end = ft
                    else:
                        unix_end = "31.12." + str(int(st[6:]) + year)

                    delta_unix = calculate_delta(unix_st, unix_end)
                    self.working_dates_start.append(unix_st)
                    self.working_dates_end.append(unix_end)
                    intervals_years.append(delta_unix)

                    if counter == 0:
                        Nr_clone = self.Nr[index_Nr]
                        self.Nr.insert(index_Nr, Nr_clone)

                        ids_clone = ids[index_Nr]
                        ids.insert(index_Nr, ids_clone)

                        repeted_wh_ids.append(index_Nr)

                        index_Nr = index_Nr + 1
                        counter = counter + 1

                self.intervals.append(intervals_years)
            index_Nr = index_Nr + 1
            index += 1

    def generate_fix_workers(self, lista_datas, ids, first_year, last_year, Nr, entity, df, pre_define_workers):

        h = []
        new_Nr = []
        new_ids = []
        index_wh = 0
        worker_zero = worker.Worker(0, 0, 0)
        worker_pre_list = []
        data_start_pre = []
        data_end_pre = []
        aps_distributed = []

        for idx, (start_date, end_date, interval_hours, pre_wk) in enumerate(
                zip(self.dates_st, self.dates_ft, self.hours, pre_define_workers)):
            interval_hours = float(interval_hours)

            # Calculate total months between start and end dates
            total_months = months_between(start_date, end_date)

            # Calculate average hours per month for this interval
            avg_hours_per_month = interval_hours / total_months

            checker = 0

            if pre_wk[0] != 0:
                workers_pre = [item for w_pp in pre_wk for item in w_pp.split(";")]
                for w_p in workers_pre:
                    w_s = w_p.split(" ")
                    if w_s[0] == entity:
                        new_Nr.append(Nr[index_wh])
                        new_ids.append(ids[index_wh])

                        h.append(interval_hours)
                        for wks in worker.list_of_workers:
                            if int(w_s[1][1]) == wks.id:
                                max_m, h_s, m_s = max_consecutive_months_worker_can_work(wks,
                                                                                         datetime.datetime.strptime(
                                                                                             start_date, "%d.%m.%Y"),
                                                                                         datetime.datetime.strptime(
                                                                                             end_date, "%d.%m.%Y"),
                                                                                         first_year,
                                                                                         interval_hours)
                                if max_m == calculate_delta(start_date, end_date):
                                    update_worker(wks, avg_hours_per_month, first_year,
                                                  generate_monthly_dates(start_date, end_date))
                                    worker_pre_list.append(wks)
                                    break
                                else:
                                    worker_pre_list.append(worker_zero)
                        data_start_pre.append(start_date)
                        data_end_pre.append(end_date)
                        aps_distributed.append(ids[index_wh])
                        index_wh += 1
                        checker += 1
                if checker != 0:
                    continue

            index_wh += 1

        return worker_pre_list, data_start_pre, data_end_pre, aps_distributed

    def get_workers(self, lista_datas, ids, first_year, last_year, Nr, entity, df, pre_define_workers):
        self.workers = []
        self.working_hours = []
        self.working_dates_start = []
        self.working_dates_end = []
        self.dates_distributed = []
        worker_zero = worker.Worker(0, 0, 0)

        h = []
        new_Nr = []
        new_ids = []
        index_wh = 0
        list_pre_def = []

        ids = list(dict.fromkeys(ids))  # Remove duplicates
        Nr = list(dict.fromkeys(Nr))  # Remove duplicates

        # calculating pre define workers
        worker_pre_list, data_start_pre, data_end_pre, aps_distributed = self.generate_fix_workers(lista_datas, ids,
                                                                                                  first_year,
                                                                                                  last_year, Nr,
                                                                                                  entity, df,
                                                                                                  pre_define_workers)
        index_pre = 0

        for idx, (start_date, end_date, interval_hours,pre_wk) in enumerate(
                zip(self.dates_st, self.dates_ft, self.hours,pre_define_workers)):
            interval_hours = float(interval_hours)

            # Calculate total months between start and end dates
            total_months = months_between(start_date, end_date)

            # Calculate average hours per month for this interval
            avg_hours_per_month = interval_hours / total_months

            if float(interval_hours) <= float(0):
                index_wh += 1
                continue

            list_ent = []

            if pre_wk[0] != 0:
                for strs in pre_wk:
                    list_ent.append(strs.split(" ")[0])
                if entity in list_ent:
                    self.workers.append(worker_pre_list[index_pre])
                    self.working_dates_start.append(data_start_pre[index_pre])
                    self.working_dates_end.append(data_end_pre[index_pre])
                    self.aps_number_distributed.append(aps_distributed[index_pre])
                    new_Nr.append(Nr[index_wh])
                    new_ids.append(ids[index_wh])
                    h.append(interval_hours)
                    index_pre += 1
                    index_wh += 1
                    list_pre_def.append(1)
                    continue

            ch_workers, hours, dates = choose_workers(start_date, end_date, interval_hours, first_year, last_year)

            if len(ch_workers) == 1:
                new_Nr.append(Nr[index_wh])
                new_ids.append(ids[index_wh])

                h.append(hours[0])
                self.workers.append(ch_workers[0])
                self.working_dates_start.append(start_date)
                self.working_dates_end.append(end_date)
                self.dates_distributed.append(dates)
                self.aps_number_distributed.append(ids[index_wh])
                list_pre_def.append(0)

            else:
                workers_array = np.zeros((len(worker.list_of_workers) + 1))

                for wk, wh in zip(ch_workers, hours):
                    workers_array[wk.id] += wh
                    list_pre_def.append(0)

                for index, hours_worked in enumerate(workers_array):
                    if hours_worked > 0:
                        new_Nr.append(Nr[index_wh])
                        new_ids.append(ids[index_wh])
                        h.append(hours_worked)
                        if index == 0:
                            self.workers.append(worker.Worker(0, 0, 0))
                        else:
                            self.workers.append(worker.list_of_workers[index - 1])
                        self.working_dates_start.append(start_date)
                        self.working_dates_end.append(end_date)

            index_wh += 1
        return h, new_ids, new_Nr, list_pre_def


def max_consecutive_months_worker_can_work(w, start_date, end_date, first_year, required_hours, av_h_min=0.0):
    current_date = start_date
    total_available_hours = w.hours_available
    consecutive_months = 0
    months = []
    st = start_date.strftime("%d.%m.%Y")
    ft = end_date.strftime("%d.%m.%Y")
    delta = calculate_delta(st, ft)
    if av_h_min == 0:
        av_hours = required_hours / delta
    else:
        av_hours = av_h_min

    times_in_the_year = 0
    year_before = current_date.year
    first_time = False
    while current_date <= end_date:
        month = current_date.month
        year = current_date.year
        available_monthly_hours = w.hours_available_per_month[year - first_year][month - 1]
        starting_year = current_date.year
        if starting_year != year_before:
            times_in_the_year = 1
            year_before = current_date.year
        else:
            times_in_the_year += 1

        if available_monthly_hours >= av_hours and total_available_hours[
            year - first_year] >= av_hours * times_in_the_year:
            consecutive_months += 1
            months.append(current_date)
        else:
            break

        # Move to the next month
        current_date = (current_date.replace(day=1) + datetime.timedelta(days=32)).replace(day=1)

    return consecutive_months, av_hours, months


def get_min_wh(w, current_date, finishing_date, first_year):
    minimum_wh = 1000

    while current_date < finishing_date:
        year = current_date.year
        month = current_date.month

        if w.hours_available_per_month[year - first_year, month - 1] < minimum_wh:
            minimum_wh = w.hours_available_per_month[year - first_year, month - 1]
        current_date = (current_date.replace(day=1) + datetime.timedelta(days=32)).replace(day=1)

    minimum_wh = round_0_25(minimum_wh)

    return minimum_wh


def choose_workers(start_date, end_date, required_hours, first_year, last_year):
    current_date = datetime.datetime.strptime(start_date, "%d.%m.%Y")
    finishing_date = datetime.datetime.strptime(end_date, "%d.%m.%Y")
    remaining_hours = required_hours
    work_distribution = []
    hours_distribution = []
    dates_distribution = []
    total_months = calculate_delta(start_date, end_date)
    av_wh = required_hours / total_months
    loop = False
    locked = 0
    counter = 0
    worker_zero = worker.Worker(0, 0, 0)

    while remaining_hours > 0 and current_date <= finishing_date:

        if locked == 0:
            counter += 1
        locked = 0

        for w in worker.list_of_workers:

            av_wh = remaining_hours / total_months

            if loop:
                av_wh = get_min_wh(w, current_date, finishing_date, first_year)

            if remaining_hours <= 0:
                break

            # how many months can this worker work
            max_months, av_wh, dates = max_consecutive_months_worker_can_work(w, current_date, finishing_date,
                                                                              first_year,
                                                                              remaining_hours, av_wh)
            if max_months > 0 and round(av_wh * max_months, 4) >= 0.25:
                wh = round_0_25(round(av_wh * max_months, 7))
                work_distribution.append(w)
                if remaining_hours - round(av_wh * max_months, 4) == 0:
                    remaining_hours -= round(av_wh * max_months, 4)
                    hours_distribution.append(round(av_wh * max_months, 4))
                else:
                    remaining_hours -= wh
                    hours_distribution.append(wh)
                dates_distribution.append([dates])
                update_worker(w, round(round_0_25(av_wh * max_months) / max_months, 7), first_year, dates)
                locked += 1

        if counter > 2:
            work_distribution.append(worker_zero)
            hours_distribution.append(round(remaining_hours, 7))
            remaining_hours = 0
        loop = True

    if len(work_distribution) == 0:
        work_distribution.append(worker_zero)
        hours_distribution.append(required_hours)

    return work_distribution, hours_distribution, dates_distribution


def generate_monthly_dates(start_date_str, end_date_str):
    # Parse input date strings to datetime.date objects
    start_date = datetime.datetime.strptime(start_date_str, '%d.%m.%Y').date()
    end_date = datetime.datetime.strptime(end_date_str, '%d.%m.%Y').date()

    # List to store the monthly dates
    monthly_dates = []

    # Start at the first day of the month of the start date
    current_date = start_date.replace(day=1)

    # Loop until the current_date exceeds the end_date
    while current_date <= end_date:
        # Append the current month date to the list
        monthly_dates.append(current_date)
        # Move to the first day of the next month
        next_month = current_date.month % 12 + 1
        next_year = current_date.year + (current_date.month // 12)
        current_date = current_date.replace(year=next_year, month=next_month, day=1)

    return monthly_dates


def update_worker(w, avg_hours, first_year, dates):
    for d in dates:
        w.hours_available[d.year - first_year] -= round(avg_hours, 4)
        w.hours_available_per_month[d.year - first_year][d.month - 1] -= round(avg_hours, 4)


def calculate_delta(st, ft):
    # Calculate the delta (duration) between two dates in months
    unix_st = datetime.datetime.strptime(st, "%d.%m.%Y")
    unix_st = int(unix_st.timestamp())

    unix_end = datetime.datetime.strptime(ft, "%d.%m.%Y")
    unix_end = int(unix_end.timestamp())

    delta_unix = unix_end - unix_st
    delta_unix = round(delta_unix / 60 / 60 / 24 / 30)

    return delta_unix


def round_0_25(duration):
    comparator = 0
    while comparator + 0.25 <= duration:
        comparator += 0.25
    return comparator
