# a simple script that will write a number into active program window
# user enters some number representing minutes of run of the program
# printed number equals to active minute

import time
from pynput.keyboard import Key, Controller

keyboard = Controller()

def typing():

     # number of iterations
    counter = 1
    
    # define delay between each button press, units: seconds
    delay = 60

    # ask user how long script operates in minutes
    wait = input('enter time in minutes: ')

    # visual confirmation of user's choince
    print('you have entered: ' + str(wait))

    # convert user entered value to integers
    wait = int(wait)

    # as long as user defined value is bigger than actual iteration:
    while wait > counter:

        # write actual iteration for visual checking 
        keyboard.type(str(counter) + ' ')

        # wait defined period of time
        time.sleep(delay)

        # incerease counter value
        counter = counter +1

# call the program        
typing()
