from datetime import datetime

while True:
    date_str = input("Enter a date (YYYY-MM-DD): ")
    try:
        # Attempt to parse the string into a date object
        date_object = datetime.strptime(date_str, "%Y%m%d").date()
        print(f"You entered: {date_object}")
        break  # Exit the loop if parsing is successful
    except ValueError:
        print("Invalid date format. Please use YYYY-MM-DD.")
