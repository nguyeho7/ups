from anthill import anthill, anthill_card, anthill_name, anthill_list
from inout_gsheet import auth_log, try_this_month, update_worksheet, update_worksheet_in, user_sheet_log, wks_in_time_log, in_log_batch
import time
from evdev import InputDevice, ecodes
from multiprocessing import Process, Value


def setup_inout(file_name='/github/ups_anthill/ups_modul/inout/anthill_inout_log.txt'):
    '''
    setup of the script with parsing the log
    '''
    parse_time_format = "%W %j %A %d.%m.%Y %H:%M:%S"
    with open(file_name) as fn:
        for row, log_txt in enumerate(fn):
            parse_line = log_txt.strip().split(" ")
            parse_name = parse_line[0]
            parse_status = parse_line[1]
            time_log = " ".join(parse_line[2:])
            parse_time_raw = time.strptime(time_log, parse_time_format)
            setup_user(anthill_name, parse_name, parse_status, parse_time_raw)

def setup_user(anthill_name, parse_name, parse_status, parse_time_raw):
    '''
    update user information from log for first run of the script
    '''
    if parse_name in anthill_name:
        anthill_name[parse_name].status = parse_status
        anthill_name[parse_name].time = parse_time_raw
        if parse_status == "IN":
            anthill_name[parse_name].day_in = parse_time_raw[2]

def parse_backup_list(file_name='/github/ups_anthill/ups_modul/inout/anthill_list.txt'):
    '''
    parse the old list that is created after each iteration with the input ID
    '''
    old_list = []
    with open(file_name) as fn:
        for row, log_txt in enumerate(fn):
            parse_line = log_txt.strip().split(" ")
            parse_name = parse_line[0]
            old_list.append(parse_name)
    return old_list

def compare_lists(old_list, anthill_list):
    '''
    compares two lists and returns a list of tuples
    each tuple contains (name, "removed"/"added", it's index)
    '''
    changes = []
    set_l1 = set(old_list)
    set_l2 = set(anthill_list)
    # ^ .. symmetric set difference
    diff = set_l1 ^ set_l2
    for x in diff:
        if x in old_list:
            change_location = old_list.index(x)
            changes.append((x, change_location, "removed"))
        elif x in anthill_list:
            change_location = anthill_list.index(x)
            changes.append((x, change_location, "added"))
    changes.sort(key = lambda change: change[1])
    return changes

def delete_row(changes, wks):
    '''
    delete the old users rows
    '''
    for (x,y,z) in reversed(changes):
        if z == "removed":
            row_detele(y, wks)

def insert_row(changes, wks):
    '''
    add new users rows
    '''
    for (x,y,z) in changes:
        if z == "added":
            row_insert(y, wks)

def check_new_users():
    '''
    check the list whether new user was added
    the whole checking new users and deleting and adding
    '''
    old_list = parse_backup_list()
    changes = compare_lists(old_list, anthill_list)
    wks, this_month, days = try_this_month()
    delete_row(changes, wks)
    insert_row(changes, wks)
    update_worksheet(anthill, anthill_name, this_month, days, wks)
    update_worksheet_in(anthill, this_month, days)

def get_next_card():
    '''
    reading the user ID
    '''
    device = InputDevice("/dev/input/event0")
    input_id = ""
    for event in device.read_loop():
        if event.type == ecodes.EV_KEY:
            input_id += str(event.code)
        if input_id.endswith("2828"):
            return input_id

def timer(anthill_card, input_id):
    '''
    timer between inout, return t_delta
    '''
    user_time = anthill_card[input_id].time
    t = time.mktime(time.localtime())
    t_new = time.mktime(user_time)
    t_delta = t - t_new
    t_new = time.localtime()
    anthill_card[input_id].time = t_new
    return t_delta

def inout(anthill_card, anthill_name, input_id, t_delta, this_month, wks):
    '''
    inout switch and logging
    '''
    user_name = anthill_card[input_id].name
    hours_delta = float(t_delta / 3600)
    if anthill_card[input_id].status == 'OUT':
        day = time.localtime()
        anthill_card[input_id].day_in = day.tm_mday
        day_in = anthill_card[input_id].day_in
        anthill_card[input_id].status = 'IN'
        wks_in_time_log(anthill_name, this_month, day_in, user_name)
    elif anthill_card[input_id].status == 'IN':
        day_in = anthill_card[input_id].day_in
        user_sheet_log(anthill_name, day_in, user_name, hours_delta, wks)
        anthill_card[input_id].status = 'OUT'

def in_log(anthill_name):
    '''
    logging all user with IN status into google sheet "Anthill IN"
    '''
    in_list = []
    time_format = "%H:%M"
    for keys, values in anthill_name.items():
        if values.status == "IN":
            time_sheet = time.strftime(time_format, values.time)
            in_list.extend((values.name, values.status, time_sheet))
    print(in_list)
    wks = auth_log("Anthill IN")
    in_log_batch(in_list, wks)

def log_txt(anthill_card, input_id):
    '''
    logging the user data to txt log
    '''
    time_format = "%W %j %A %d.%m.%Y %H:%M:%S"
    user_name = anthill_card[input_id].name
    user_status = anthill_card[input_id].status
    user_time = anthill_card[input_id].time
    f = open('/github/ups_anthill/ups_modul/inout/anthill_inout_log.txt', "a")
    f.write("{} {} {}\n".format(user_name, user_status , time.strftime(time_format, user_time)))
    f.close()

def backup_list(anthill_list):
    '''
    write a list of current anthill empployyes after each ittereation with input ID
    '''
    f = open('/github/ups_anthill/ups_modul/inout/anthill_list.txt', "w")
    for x in range(len(anthill_list)):
        user_name = anthill_list[x]
        f.write("{}\n".format(user_name))
    f.close()

# main loop waiting for an user ID. updating user data
def loop():
    while True:
        # input_id = input("kdo jsi?:")
        input_id = get_next_card()
        print(input_id) #print the ID to console. for adding new user
        if input_id in anthill_card:
            wks, this_month, days = try_this_month()
            t_delta = timer(anthill_card, input_id)
            inout(anthill_card, anthill_name, input_id, t_delta, this_month, wks)
            in_log(anthill_name)
            log_txt(anthill_card, input_id)
            backup_list(anthill_list)
        else:
            continue

# weird stuff
if __name__ == "__main__":
    setup_inout()
    connected = False
    while connected == False:
        try:
            check_new_users()
            connected = True
        except:
            time.sleep(10)
    in_log(anthill_name)
    loop()
