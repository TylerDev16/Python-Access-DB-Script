from imaplib import Int2AP
import time
import pyodbc
from art import *
from halo import Halo


tprint("Inventory Master 5000")
print("-TT \n")
spinner = Halo(text='Loading')
spinner.start()                    
time.sleep(1)
spinner.stop()

#connect to database
print("Which database do you want to connect to?")
db_choice = input("1. Live DB on //FS1US/BNA \n2. Test DB on //FS1US/BNA.\n")
if db_choice == "1":
    conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=\\fs1us\bna\Departments\Information Technology\Database\ITDatabase.accdb')
elif db_choice == "2":
    conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=\\fs1us\bna\Departments\Information Technology\Database\TestITDatabase.accdb')
cursor = conn.cursor()

while True:

    print("Please select from an option below")
    choice = input("1. Add New Equipment. \n2. Add New Equipment Type. \n3. Add New Model Type. \n4. Assign Equipment Co Asset No\nPress 'q' to exit script \n")
    print("--------------------------")

    if choice == "1":
        
        # get equipment type from db
        cursor.execute("SELECT EquipmentType FROM EquipmentType \n")
        et_row = [item[0] for item in cursor.fetchall()]
        for (i, item) in enumerate(et_row, start=0):
            print(i, item)
        
        # assign user input for et to variable
        et_choice = int(input("Select # from Equipment Type list. \n"))
        equipment_type = et_row[et_choice]
        print(equipment_type)

        # if user selects "Desktop" show only desktop models.
        if et_choice == 0:

            cursor.execute("SELECT ModelType FROM ModelTypes WHERE EquipmentType = 'Desktop'  \n")
            mt_row = [item[0] for item in cursor.fetchall()]
            for (i, item) in enumerate(mt_row, start=0):
                print(i, item)

            # assign user input for mt to variable
            mt_choice = int(input("Select # from Equipment Model list. \n"))
            equipment_model = mt_row[mt_choice]
            print(equipment_model)
        
        # if user selects "Laptop" show only laptop models.
        elif et_choice == 3:
            
            cursor.execute("SELECT ModelType FROM ModelTypes WHERE EquipmentType = 'Laptop'  \n")
            mt_row = [item[0] for item in cursor.fetchall()]
            for (i, item) in enumerate(mt_row, start=0):
                print(i, item)

            # assign user input for mt to variable
            mt_choice = int(input("Select # from Equipment Model list. \n"))
            equipment_model = mt_row[mt_choice]
            print(equipment_model)

        # if user selects "Monitor" show only monitor models.
        elif et_choice == 4:
            
            cursor.execute("SELECT ModelType FROM ModelTypes WHERE EquipmentType = 'Monitor'  \n")
            mt_row = [item[0] for item in cursor.fetchall()]
            for (i, item) in enumerate(mt_row, start=0):
                print(i, item)
            
            # assign user input for mt to variable
            mt_choice = int(input("Select # from Equipment Model list. \n"))
            equipment_model = mt_row[mt_choice]
            print(equipment_model)
            
        else:        
            cursor.execute("SELECT ModelType FROM ModelTypes WHERE EquipmentType NOT IN ('Monitor', 'Laptop', 'Desktop')  \n")
            mt_row = [item[0] for item in cursor.fetchall()]
            for (i, item) in enumerate(mt_row, start=0):
                print(i, item)
        
            # assign user input for mt to variable
            mt_choice = int(input("Select # from Equipment Model list. \n"))
            equipment_model = mt_row[mt_choice]
            print(equipment_model)

        print("--------------------------")

        # assign Co Asset Number
        co_num_input = input("What is the Co Asset Number? (Leave blank if none) \n")
        if co_num_input == "":
            co_num_input = "000000"

        # get lease bool
        print("------------------------")
        lease = input("Is this a lease? Y/N\n")
        if lease == "Y":
            lease = True
        elif lease == "N":
            lease = False

        # retrieve all serial numbers currently in db
        cursor.execute ("SELECT SN FROM Equipment")
        sn_row = [item[0] for item in cursor.fetchall()]

        # get serial numbers for equipment
        print("------------------------")
        serial_numbers = []
        print("Scan barcode, Type 'e' to delete last scanned item. Type 'q' when you are done.")
        # loop for user to insert multiple serial numbers.
        while True:
            s = input("Scan Code: ")
            # add scanned sn to list
            if s == "e":
                serial_numbers.pop()
                print(serial_numbers)
                continue
            # check if sn exists already in db
            elif s in sn_row:
                print("Serial # already exists in DB. Try again.")
                continue
            elif s == "q":
                break
            serial_numbers.append(s)
        print("Here is the list of serial numbers:")
        print(serial_numbers)
        x = len(serial_numbers)
        print (x)

        # insert data into db
        for i in serial_numbers:
            cursor.execute('''
                        INSERT INTO Equipment (SN, [Co Asset No], EquipmentModel, Lease)
                        VALUES (?,?,?,?)
                        ''', (i, co_num_input , equipment_model, lease))

        print("------------------------")

        print("Inserting into database.")
        conn.commit()
        time.sleep(2)
        print("Success! \n")



    if choice == "2":
        et_name = input("What is equipment type name? \n")
        cursor.execute("""INSERT INTO EquipmentType (EquipmentType)
                        VALUES (?) """, (et_name))
        print("Success! '{}', Added to database \n".format(et_name))
        conn.commit()


    if choice == "3":
        mt_name = input("What is model type name? \n")
        mt_manuf = input("Who is the manufacturer? \n")
        mt_model_num = input("Scan Model Number Barcode: \n")
        cursor.execute("""INSERT INTO ModelTypes (ModelType, Manufacturer, [Model Number])
                        VALUES (?,?,?) """, (mt_name, mt_manuf, mt_model_num))
        print("Success! '{}', Added to database \n".format(mt_name))
        conn.commit()

    if choice == "4":

        sn_search = input("Scan barcode to search equipment: \n")
        cursor.execute ("SELECT (SN) FROM Equipment WHERE SN=?", sn_search)
        sn_row = [item[0] for item in cursor.fetchall()]
        print(sn_row)

        cursor.execute("SELECT [Co Asset No], EquipmentModel FROM Equipment WHERE SN=?", sn_search)
        sn_co_asset_num = [item[0] for item in cursor.fetchall()]
        print(sn_co_asset_num)

        cursor.execute("SELECT EquipmentModel FROM Equipment WHERE SN=?", sn_search)
        sn_equip_model = [item[0] for item in cursor.fetchall()]
        print(sn_equip_model)

        user_select = input("Would you like to assign this equipment a number? Y/N \n")

        if user_select == "Y":
            sn_co_assign = input("What number to be assigned to equipment? (000000)\n")
        
            print("Inserting into database...")
            time.sleep(1)
            cursor.execute("UPDATE Equipment SET [Co Asset No] = ? WHERE SN = ?", (sn_co_assign, sn_search))
            conn.commit()
            print("success")
        
        if user_select == "N":
            break
        

    if choice == "q":
        break
        