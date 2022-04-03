import itertools
from string import digits, punctuation, ascii_letters
import win32com.client as client
from datetime import datetime
import time

print('Version 1.0 \nMade by Fengwuu')

def brute_excel_doc():
    print('****Hello friend!****')

    try:
        dir_file = input('Enter file directory')
        password_length = input('Enter the approximate length of the password, for example 6-7: ')
        password_length = [int(item) for item in password_length.split("-")]
    except:
        print('Check input')

    print('If password contains digits, press : 1\nIf password contains letters, press: 2\n'
          'If password contains digits and letters,press: 3\nIf password contains digits,letters and symbols,press: 4')

    try:
        choice = int(input(": "))
        if choice == 1:
            possible_symbols = digits
        elif choice == 2:
            possible_symbols = ascii_letters
        elif choice == 3:
            possible_symbols = digits + ascii_letters
        elif choice == 4:
            possible_symbols = digits + ascii_letters + punctuation
        else:
            possible_symbols = 'Check input'
    except:
        print("Check input")

    start_timestamp = time.time()
    count = 0
    for pass_length in range(password_length[0], password_length[1] + 1):  # starting
        for password in itertools.product(possible_symbols, repeat=pass_length):
            password = "".join(password)
            opened_doc = client.Dispatch('Excel.Application')
            count += 1
            try:
                opened_doc.Workbooks.Open(  # opening file
                    dir_file, False, True, None, password)
                time.sleep(0.1)
                print(f"Finished at - {datetime.utcfromtimestamp(time.time()).strftime('%H:%M:%S')}")
                print(f'Password cracking time - {time.time() - start_timestamp}')
                return f'Attempt #{count} Password is: {password}'
            except:
                print(f'Attempt #{count} Incorrect password: {password}')
                pass


def main():
    print(brute_excel_doc())


if __name__ == '__main__':
    main()

print("Bye")
time.sleep(50)

