from program_file.merging_two_files import merging_two_files

while True:
    print(' Menu '.center(48, '='))
    
    choise = input('\n1. merging import and export files\n2. exit\n\nYou choosed: ')
 
    if choise == '1':
        merging_two_files()
    elif choise == '2':
        break
