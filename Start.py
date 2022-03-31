from openpyxl import Workbook, load_workbook

# olika variablar
wb = load_workbook("Grupparbete.xlsx")
ws = wb.active
ws_map = wb["Map"]
ws_info = wb["Rooms"]
ws_save = wb['Save']
ID = ws_map["A1"]
# sätter playpos som startrum
player_position = ID
x = 1
y = 2
# dessa 2 gör look variable
room_num = int(player_position.value)
room_id = room_num + 1
direct_check = True


# skapar x och y i en "karta"
def playerpos(m_a):
    map = []
    for y in m_a:
        x_rooms = []
        for x in y:
            x_rooms.append(str(x.value))
            map.append(x_rooms)
    return map


def save():
    global player_position
    player_position = ws_save['A1']


# printar ut info och loopar i worksheet save där rummen ligger som du har besökt som save har sparat
def info():
    for x in range(30):
        print(wsSave['A1'])


# Testkör värden för att se att inladdning gått rätt till såhär långt

# Skapar en lista med all information från rooms
roomList = []

rumJagHarBesokt = {}
rumJagHarBesokt.append(ID)

for x in range(len(rumJagHarBesokt)):
    print(rumJagHarBesokt[x])


# printar desc av rummet man är i
def look():
    print(ws_info[f"C{room_id}"].value)


# move funktionen
def move(direction):
    # gör en funktion som kan ändra plats/riktning i move dvs att x och y ändras vid en flyttning
    # funktionen baseras på global därför så gör jag x och y global
    global x
    global y
    global direct_check
    # kollar så att den riktningen finns med hjälp av variabel som ändras
    direct_check = True

    if direction == "north":
        y += 1
        # har innan nolställt x och gett y värdet 1, därför om man går norr så ändras y variabeln +1
    elif direction == "south":
        y += -1
        # samma här men olika värden^
    elif direction == "west":
        x += -1
        # samma här men x ändras, notera att väst och öst är x
    elif direction == "east":
        x += 1

    else:
        # ifall riktningen inte finns så kan man ej move åt den riktningen
        print("No direction")
        direct_check = False
        # om inget av de ovanför runnar av en viss anledning så printar error
    return x, y


# vi måste självklart få en returnering när koden runnar^
# while loop med ågra varaiblar som player pos och kartan
def start_up():
    global room_id
    global player_position
    player_position = ws_save['A1']
    start = True
    map = playerpos(ws_map)
    player_position = map[y][x]
    # input som söker konstant efter olika ord och sedan händer det som skrivs
    while start:
        reading = input("")
        if reading == "look":
            look()
        elif reading == "stop":
            start = False
        elif reading == "move":
            move(direction=input("what direction".lower()))
            # checkar ifall en sådan riktning fanns
            if direct_check:
                # printar x och y värde samt cellvärde
                print(x, y)
                print(player_position)
                # dessa 2 uppdaterar room number efter man har flyttat sig
                room_num = float(player_position)
                room_id = int(room_num + 1)
                # win eller lose ifall:
                if room_id == 7:
                    print("you won")
                    start = False
                elif room_id == 8:
                    print("you lost")
                    start = False
        # savear
        elif reading == "save":
            save()
        # felmedelande
        else:
            print("No such thing or you miss spelled")
        # uppdaterar mapens x och y
        player_position = map[y][x]


start_up()
