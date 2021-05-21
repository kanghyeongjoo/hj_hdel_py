import re

def special_str_split(floor_and_floor_height):
    comma_split = {}
    tilde_split = []
    for before_floor, height in floor_and_floor_height.items():
        before_floor=before_floor.replace(" ","")
        if "," not in before_floor and "." not in before_floor:
            comma_split.update({before_floor: height})
        elif "," in before_floor or "." in before_floor:
            comma_split_list = re.split("[,.]", before_floor)
            for split_floor in comma_split_list:
                comma_split.update({split_floor: height})

    for before_floor, height in comma_split.items():
        if "~" not in before_floor and "-" not in before_floor:
            tilde_split.append([before_floor, height])
        elif "~" in before_floor or "-" in before_floor:
            str_floor = re.findall("(\w+)\W", before_floor)[0]
            str_text = re.findall("(\D+)\d+", str_floor)
            end_floor = re.findall("\W(\w+)", before_floor)[0]
            end_text = re.findall("(\D+)\d+", end_floor)
            if len(str_text) == 0: # start 층표기에 B2~3과 같은 문자가 있는지 확인
                st_no = int(str_floor)
                end_no = int(end_floor) + 1
                for floor in range(st_no, end_no):
                    tilde_split.append([str(floor), height])
            elif len(end_text) == 0: #start 층표기에는 문자가 있고, end 층표기에는 문자가 없을 때
                text = str_text[0]
                st_no = re.findall("\d+", str_floor)[0]
                st_no = int(st_no)
                end_no = int(end_floor) + 1
                for floor in range(st_no, 0, -1):
                    tilde_split.append([text+str(floor), height])
                for floor in range(1, end_no):
                    tilde_split.append([str(floor), height])
            elif len(end_text) > 0: #start, end 모두 층표기에 문자가 있을 때
                text = str_text[0]
                st_no = re.findall("\d+", str_floor)[0]
                st_no = int(st_no)
                end_no = re.findall("\d+", end_floor)[0]
                end_no = int(end_no) - 1
                for floor in range(st_no, end_no, -1):
                    tilde_split.append([text+str(floor), height])

    return tilde_split
