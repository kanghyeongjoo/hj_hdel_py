import re
litt={"이름":"가가", "나이":11, "성별":"여성"}
if "나이" in litt:
    print("key ok")
if "가가" in litt.values():
    print("value ok")

print("50%".replace("%", ""))

def ttt():
    t="700%"
    if t == "700%":
        af_t, at_t = tty(t)
    print(af_t, at_t)

def tty(textstring):
    textstring = textstring.rstrip("%")
    attaa = "확인중"
    return attaa, textstring
ttt()

print("dfd5232%".rstrip('%'))

dd = "'3ø 4선 380V / 1ø 220V 60HZ"
dd=dd.split(" ")
print(dd)


for ii in range(0,4):
    print("forforfor")


def ddd():
    lk={}
    for ill in range(2):
        af_t, bt_t = ddy(ill)
        for kk in range(0,len(af_t)):
            lk.update({af_t[kk]:bt_t[kk]})
    print(lk)

def ddy(ss):

    if ss == 0:
        attaa = ["발란스"]
        tedd = ["50"]
    elif ss == 1:
        attaa = ["동력", "접지", "주파수"]
        tedd = ["380", "220", "60"]

    return attaa, tedd

ddd()


kd=[]
attaa = ["동력", "접지", "주파수"]
for si in attaa:
    kd.append(si)
print(kd)

ad=["'3ø", '4선', '380V', '/', '1ø', '220V', '60HZ']
for j in ad:
    if "V" in j or "HZ" in j:
        print(j)

slqk="1 car 2 buttons"
lldfdf=[" ","car"]
dkdk = ["2", "bc"]

dictionary = {ord(" ") : "", "car": "C", "2": "BC"}

dihnk = slqk.translate(dictionary)
print("Dfdfdfdfdfdfd : ", dihnk)


sldf=slqk.replace(" ", "").replace("CAR", "C")
print(sldf)

fdfdf=slqk.upper().replace(" ", "").replace("CAR", "C").replace("BUTTONS","BC")
print("fdfdfd", fdfdf)

p1p1='VVVF(WBSS)'
if "WBSS" in p1p1:
    print("p1p1")


spd = "150M/min"
spdid = spd.lower().find("m")
spdd  = spd[:spd.lower().find("m")]
print(spdd)

drrr = "1 CAR 2BC"

drrd = re.findall("\d+", drrr)
print(drrd)

ckj="".join(["C", "dk"])
print(ckj)

popoo='3ø 4선 380V / 1ø 220V 60HZ'

poo = re.findall("\d\d\dV|\d\dHZ", popoo)
print(poo)

dj = ["1", "2", "3"]
trs_tag=[]
for no in range(len(dj)):
    trs_tag.append("@NO" + str(no))

print(trs_tag)


ee = "인화물용(장애인기능)"

eeee = re.findall("\w\W", ee)
aeee = re.finditer("\W", ee)
for ae in aeee:
    see = ae.span()
    print(see)

print(eeee)
print(aeee)

el = "인화,장애"
eeel = re.findall("\w\W", el)
aeel = re.finditer("\w\W", el)
for ael in aeel:
    see = ael.span()
    print(see)


print(eeel)
print(aeel)

el = "병원,장애,라라"
eeel = re.findall("\W", el)
aeel = re.finditer("\W", el)
ankn
for ael in aeel:
    zzz = ael.start()
    see = ael.end()
    print(zzz)
    print(see)
    print("rererererererererererere", el[:zzz].pop)

print(eeel)
print(aeel)




edl = "병애"
eedel = re.findall("\W", edl)
print(len(eedel))

print("Dcvdddddddddddddddddvvvvvv", edl[:2])
print("dkdkdk", eedel)



e2222l = "병원,장애,라라"
eeel = re.findall("\W", e2222l)
aeel = re.finditer("\W", e2222l)
ankn = []

for ael in aeel:
#    print(ael.span())
    zzz = ael.start()
    see = ael.end()
    print(zzz)
    print(see)
    kkk = e2222l[:ael.start()]
    ankn.append(kkk)
    e2222l = e2222l.lstrip(kkk)

print(ankn)

e2 = "병원,장애,라라"
lll = []

while re.search("\W", e2) != None :
    print(re.search("\W", e2))
    zst = re.search("\W", e2).start()
    zed = re.search("\W", e2).end()
    lll.append(e2[:zst])
    e2 = e2.lstrip(e2[:zed])

print("lll ", lll)

textstring = "비상,관통"
trs_textstring=[]
while re.search("\W", textstring) != None:
    spc_st = re.search("\W", textstring).start()
    spc_ed = re.search("\W", textstring).end()

    trs_textstring.append(textstring[:spc_st][:2])
    textstring = textstring.lstrip(textstring[:spc_ed])

print("eee", trs_textstring)


dfdfd = "380V / 1ø 220V 60HZ"

fkfkf = re.sub("\s", "", dfdfd)
bf_text = re.findall("(\d\d\d)V|(\d\d)HZ", fkfkf)

print(bf_text)
