import string

test0={'1': '4020', '2': '2870', '3~8': '2900', '9': '4350'}

test1={'1': '4020', '2': '2870', '3,4,5,6,7,8': '2900', '9': '4350'} #콤마 형식은 split으로 분리하기

#물결과 콤마, 콤마와 물결 형식으로 작성하는 경우도 고려해야해(물결, 콤마 인덱스 번호를 비교해서 먼저 있는 것부터 해결해나간다)
#뒤에 있는 경우 뒤에 있는 인덱스까지를 추출
#숫자가 아닌 경우 뺄셈이 불가하므로 추가 고려 필요
test2={'1': '4020', '2': '2870', '3~7,R': '2900', '9': '4350'}

test3={'1': '4020', '2': '2870', 'F,6~8': '2900', '9': '4350'}

test3={'1': '4020', '2': '2870', '6~8': '2900', '9~12': '4350'}

test10={'1': '4020', '2': '2870', '13~18': '2900', '9': '4350'}



floor_comma_mark= {}
for be_floor_mark, hight in test10.items():
    if "," not in be_floor_mark:
        floor_comma_mark.update({be_floor_mark:hight})
    elif "," in be_floor_mark:
        split_floor = be_floor_mark.split(",")
        for sp_floor in split_floor:
            floor_comma_mark.update({sp_floor:hight})
print(floor_comma_mark)

floor_tilde_mark={}
for be_floor_mark, hight in floor_comma_mark.items():
    if "~" not in be_floor_mark:
        floor_tilde_mark.update({be_floor_mark:hight})
    if "~" in be_floor_mark:



