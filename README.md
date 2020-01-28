# Rocket Data Aquisition Program
Program I worked on at UBC Rocket for reading sensor input from a Labjack. (won't work properly without a T series Labjack). Could be used for data aquisition during cold flow or hot fire tests. Developed with Python, and Kivy for UI.

Note: Rocket.py is the script I wrote, others are from Labjack website

## Features:
- Lots of settings for which analog pins to read from, remapping input voltage to actual value, channel names/units for output formatting,  etc.

- Read from each pin at custom frequency

- Records and saves data to excel file 

Note: Hoping to add support for Stream Mode soon (enables ultra-high frequency data collection)

