from tkinter import *
from datetime import datetime
from PIL import Image, ImageTk

import time
import sys
import xlrd


def add_label():
    global DELAY
    global countLabel
    if DELAY == 0:
        body_frame.destroy()
    if DELAY == -1:
        return
    if countLabel:
        countLabel.pack_forget()
    # print("Foward:  " + str(DELAY) + "m")
    DELAY = DELAY - 1
    start_frame.after(1000, add_label)


# Ask which service now
service_num = int(input("Which Service(1-3): "))
while service_num <= 0 | service_num >= 4:
    print("Enter between 1 - 3.")
    service_num = int(input("Which service(1-3): "))

now = datetime.now()
month_list = {1:"Jan ", 2:"Feb ", 3:"Mar ", 4:"Apr ", 5:"May ", 6:"Jun ",
              7:"Jul ", 8:"Aug ", 9:"Sep ", 10:"Oct ", 11:"Nov ", 12:"Dec "}
service_list = {1:" at 8:00 AM", 2:" at 10:30 AM", 3:" at 1:00 PM"}

month = now.month
date = str(15) # str(now.day)
year = str(now.year)

this_service = month_list[month] + date + ", " + year + service_list[service_num]  # Nov 13, 2020 at 10:30 AM
print("Today's Service: " + this_service)

loc  = ("report-2020-11-10T1105.xlsx")  # EventBrite List
loc2 = ("11082020_list.xlsx")           # pastor/volunteer

# input file prep
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

# format checking
if sheet.cell_value(0, 10-1) != "Barcode #":
    print("ERROR: Barcode # is not on excel file")
    quit()

# output file prep
logfile = (now.strftime("%Y_%m_%d__%H_%M_%S")+".log")
f = open(logfile, "x", encoding='utf-8')  # create file for writing
f.write(this_service + "\n")

# read-up ticket and makes dictionary
attendee_list = {
        "barcode#": ["name", "email", "phone", "checkin"]
}

entered_person = {
        "barcode#": "Time"
}

for i in range(sheet.nrows):
    first_name  = sheet.cell_value(i,5-1)
    last_name   = sheet.cell_value(i,6-1)
    name = first_name + " " + last_name
    email = sheet.cell_value(i,7-1)
    barcode_num = sheet.cell_value(i,10-1)
    checkIn = sheet.cell_value(i, 11 - 1)
    phone = sheet.cell_value(i,13-1)
    attendee_list[barcode_num] = [name, email, phone, checkIn]

# pastor/volunteer list loading
wb = xlrd.open_workbook(loc2)
sheet = wb.sheet_by_index(0)


for i in range(sheet.nrows):
    pName = sheet.cell_value(i,2-1)
    barcode_num = sheet.cell_value(i,1-1)
    pEmail = sheet.cell_value(i,3-1)
    pPhone = sheet.cell_value(i,4-1)
    pCheckin = this_service
    attendee_list[barcode_num] = [pName, pEmail, pPhone, pCheckin]

# print(attendee_list)

# GUIs
ROOT = Tk()
ROOT.title("EPC Sunday Service")
ROOT.tk_setPalette(background='white')
ROOT.geometry('1920x800')

wel_img1 = ImageTk.PhotoImage(Image.open("img/welcome1.png"))
wel_img2 = ImageTk.PhotoImage(Image.open("img/welcome2.png"))
wel_img3 = ImageTk.PhotoImage(Image.open("img/welcome3.png"))
qr_img = ImageTk.PhotoImage(Image.open("img/qr_instruction.png"))
confirmed_img = ImageTk.PhotoImage(Image.open("img/confirmed.png"))
notOnTime_img = ImageTk.PhotoImage(Image.open("img/not_on_time.png"))
notreg_img = ImageTk.PhotoImage(Image.open("img/not_registered.png"))
redeemed_img = ImageTk.PhotoImage(Image.open("img/redeemed.png"))

if service_num == 1:
    header = Image.open("img/welcome1.png")
if service_num == 2:
    header = Image.open("img/welcome2.png")
if service_num == 3:
    header = Image.open("img/welcome3.png")

photo = ImageTk.PhotoImage(header)
headerLabel = Label(image=photo)
headerLabel.image = photo
headerLabel.place(x=64, y=52)

# start page frame
start_frame = Frame(ROOT, padx=0, pady=0, highlightbackground='red', highlightthickness=0)
start_frame.pack()
start_frame.place(x=550, y=221)
qrLabel = Label(start_frame, image=qr_img, font="bold")
qrLabel.pack()

DELAY = 0
exit_loop = False
while exit_loop == False:
    scan_num = input("Scan QR Code: ")

    if scan_num == "q":
        exit_loop = True
    else:
        # GUI
        body_frame = Frame(ROOT, padx=0, pady=51, background='white',highlightbackground='red', highlightthickness=0)
        body_frame.pack()
        body_frame.place(x=550, y=221)
        countLabel = None

        if scan_num in attendee_list:
            now = datetime.now()
            name = attendee_list[scan_num][0]
            email = attendee_list[scan_num][1]
            phone = attendee_list[scan_num][2]
            checkIn = attendee_list[scan_num][3]

            ## REDEEMED CODE ##
            if scan_num in entered_person:
                print("ERROR: This ticket is already redeemed")
                print("       Ticket# :" + scan_num)
                print("       Name    :" + name)
                print("       Entered :" + entered_person[scan_num])
                sys.stdout.write('\r\a')
                sys.stdout.flush()
                # GUI
                Label(body_frame, image=redeemed_img, font="bold", pady=20).pack(side=TOP)
                Label(body_frame, text=name, font=("Arial", 19), pady=16).pack()
                Label(body_frame, text="Redeemed time: " + entered_person[scan_num], font=("Arial", 19)).pack()
                Label(body_frame, text=" ", font='bold', pady=16).pack()
                DELAY = 6
                add_label()

            else:
                ## RESERVATION TIME NOT MATCHED ##
                if checkIn != this_service:
                    print("Reservation time does not match.")
                    sys.stdout.write('\r\a')
                    sys.stdout.flush()
                    # GUI
                    Label(body_frame, image=notOnTime_img, font="bold", pady=20).pack(side=TOP)
                    Label(body_frame, text=name, font=("Arial", 19), pady=16).pack()
                    Label(body_frame, text="Reserved time: " + checkIn, font=("Arial", 19)).pack()
                    Label(body_frame, text=" ", font='bold', pady=16).pack()
                    DELAY = 6
                    add_label()

                ## NEW SUCCESSFUL ENTRY PROCESSING ##
                else:
                    print(  name + " " + now.strftime("%H:%M:%S"))
                    f.write(name + ", " + email + ", " + phone + ", " + now.strftime("%H:%M:%S") + "\n")
                    entered_person[scan_num] = now.strftime("%H:%M:%S")
                    # GUI
                    Label(body_frame, image=confirmed_img, font="bold", pady=20).pack(side=TOP)
                    Label(body_frame, text=name, font=("Arial", 19), pady=16).pack()
                    Label(body_frame, text="Redeemed time: " + now.strftime("%H:%M:%S"), font=("Arial", 19)).pack()
                    Label(body_frame, text=" ", font='bold', pady=16).pack()
                    DELAY = 4
                    add_label()

        ## CODE NOT EXIST ##
        else:
            print("Cannot find ticket #: " + scan_num)
            sys.stdout.write('\r\a')
            sys.stdout.flush()
            # GUI
            Label(body_frame, image=notreg_img, font="bold", pady=20).pack(side=TOP)
            Label(body_frame, text=" ", font='bold', pady=20).pack()
            Label(body_frame, text=" ", font='bold', pady=16).pack()
            Label(body_frame, text=" ", font='bold', pady=16).pack()
            DELAY = 6
            add_label()
