import random


def random_int(number, number_range: list):
    random_integers = random.sample(range(number_range[0], number_range[1]), number)
    return sorted(random_integers)


def random_server_name(server_dict, max_slot):
    server_name_list = [server_name for server_name in server_dict.keys()
                        if server_dict[server_name]['slotOccupation'] <= max_slot]

    return random.choice(list(server_name_list))
