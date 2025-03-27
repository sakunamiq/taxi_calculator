# Taxi Payment Calculator

A simple desktop application for calculating taxi driver payments based on revenue, commissions, and various bonuses.

## Features

- Calculate driver payments based on total revenue
- Automatic calculation of gasoline expenses based on revenue tiers
- Support for weekend bonuses (+5%)
- Top driver place bonuses (1st place: +5%, 2nd place: +3%, 3rd place: +1%)
- Calendar date picker for easy date selection
- Excel logging system to keep records of all calculations
- Clean, user-friendly interface

## Requirements

- Windows 10/11
- Python packages (if running from source):
  - tkinter
  - tkcalendar
  - openpyxl
  - datetime

## Installation

### Using the Executable (Recommended)

1. Download the latest release `taxi_calculator.exe` from the releases section
2. Place the executable in a dedicated folder where you want to store calculation logs
3. Double-click to run the application

### From Source

1. Clone this repository
2. Install required packages:
   ```
   pip install tkcalendar openpyxl
   ```
3. Run the application:
   ```
   python taxi_calculator.py
   ```

## Usage

1. Enter the driver's name
2. Select the date using the calendar picker (defaults to current date)
3. Select the vehicle from the dropdown
4. Enter the total revenue amount
5. Enter the commission amount
6. Check the "weekend" box if applicable
7. Select driver's ranking position if applicable
8. Click "Calculate and Save" to process the calculation
9. Results will be displayed and saved to Excel automatically

## Recent Changes

- Added calendar date picker with current date as default
- Simplified vehicle selection in dropdown
- Improved UI layout and window sizing
- Modified payment calculation display for clarity

## License

This project is licensed for personal and commercial use. 