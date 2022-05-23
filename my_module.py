STAR = '★'
ARROW = '▶'
BAR = '-'

error_brand_name = {'운송장출력', '#N/A'}
error_order_name = {'#REF!'}

def isbrand(name):
    name = name.replace(" ", "")
    if name in error_brand_name:
        return False
    return True

def isorder(name):
    name = name.replace(" ", "")
    if name in error_order_name:
        return False
    return True

# 윤년인지 확인하는 함수
def is_leap_year(year):
    if year % 4 == 0 and year % 100 != 0 or year % 400 == 0:
        return True
    else:
        return False

# 이번달의 일수를 반환하는 함수
def get_days_in_month(year, month):
    month = int(str(month).strip('0'))
    if month == 2:
        if is_leap_year(year):
            return 29
        else:
            return 28
    elif month in [1, 3, 5, 7, 8, 10, 12]:
        return 31
    else:
        return 30

def get_main_name(str, brand):
    str = remove_brand(str, brand)

    if ARROW not in str:
        if STAR not in str or BAR not in str:
            return str.strip().strip('[]')

        else:
            index_of_bar = str[::-1].index('-')
            index_of_bar = len(str)-1 - index_of_bar
            index_of_st = str[::-1].index(STAR)
            index_of_st = len(str)-1 - index_of_st
            return (str[:index_of_bar] + str[index_of_st+1:]).strip().strip('[]')
    else:
        index_of_arw = str.index('▶')
        return str[:index_of_arw].strip().strip('[]')

def get_sub_name(str):
    if '★' not in str:
            return str.strip()

    else:
        index_of_bar = str[::-1].index('-')
        index_of_bar = len(str)-1 - index_of_bar
        index_of_st = str[::-1].index(STAR)
        index_of_st = len(str)-1 - index_of_st
        return (str[:index_of_bar] + str[index_of_st+1:]).strip().replace(ARROW, " ")

def remove_brand(str, brand):
    return str.replace(brand, "").strip()

def get_sub_infos(str):
    if '▶' not in str:
        return ""
    index_of_arw = str.index('▶')
    return str[index_of_arw+1:]

def get_sub_info_list(sub_infos):
    if not sub_infos:
        return [""]
    return sub_infos.split(',')

def get_key_name(main_name: str):
    return main_name.replace(" ", "")

def get_count(name):
    if STAR not in name or BAR not in name:
        return 1
    
    index_of_st = name[::-1].index(STAR)
    index_of_st = len(name)-1 - index_of_st
    index_of_bar = name[::-1].index(BAR)
    index_of_bar = len(name)-1 - index_of_bar
    return int(name[index_of_bar+1:index_of_st-1])

def get_mainsub_name(main_name, sub_name):
    main_name = main_name.replace(ARROW, " ")
    sub_name = sub_name.replace(ARROW, " ")

    if main_name == "":
        return sub_name.strip().strip('[]')
    return (f"{main_name} {sub_name}").strip().strip('[]')

def add_to_dict(total_dict, brand, name, day, info):
    key = get_key_name(name)
    count = get_count(info)
    
    if key not in total_dict:
        total_dict[key] = [brand, name, dict()]
        total_dict[key][2][day] = count
    else:
        total_dict[key][2][day] += count

# num to excel columns (A, B, C, ...)
def num_to_excel_columns(num):
    result = ""
    while num > 0:
        num, remainder = divmod(num - 1, 26)
        result = chr(65 + remainder) + result
    return result