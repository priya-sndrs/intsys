import os
from datetime import date
from openpyxl import Workbook, load_workbook

damage_reports = []
excel_file = "damaged_items_reports.xlsx"

def save_report_to_excel(report):
    if os.path.exists(excel_file):
        wb = load_workbook(excel_file)
        sheet = wb.active
    else:
        wb = Workbook()
        sheet = wb.active
        sheet.append(["Report Number", "Item", "Chair Number", "Location", "Description", "Informant", "Date", "Status"])
    
    sheet.append([
        report["ReportNumber"],
        report["Item"],
        report.get("ChairNumber", "N/A"),
        report["Location"],
        report["Description"],
        report["Informant"],
        report["Date"],
        report["Status"]
    ])

    try:
        wb.save(excel_file)
        print(f"\nReport saved to {excel_file}\n")
    except PermissionError:
        print("\nPlease close the excel file and try again.\n")

def report_damage():
    print("\n--- Report Damaged Items ---")
    report_number = len(damage_reports) + 1
    print(f"Report Number: {report_number}\n")

    item = input("Damaged Item Name: ")
    
    chair_number = None
    if item.lower() == "chair":
        chair_number = input("Chair Number: ")
    
    location = input("Location: ")
    description = input("Damage Description: ")
    informant = input("Informant: ")
    report_date = str(date.today())

    record = {
        "ReportNumber": report_number,
        "Item": item,
        "ChairNumber": chair_number if chair_number else "N/A",
        "Location": location,
        "Description": description,
        "Informant": informant,
        "Date": report_date,
        "Status": "Pending"
    }
    damage_reports.append(record)
    print(f"\nReport Submitted Successfully! Status set to 'Pending'.")
    save_report_to_excel(record)

def view_reports():
    print("\n=== Classroom Damaged Items Report ===")
    if len(damage_reports) == 0:
        print("No reports found.\n")
    else:
        for report in damage_reports:
            print(f"\nReport {report['ReportNumber']}:")
            print(f" Item: {report['Item']}")
            if report['Item'].lower() == "chair":
                print(f" Chair Number: {report['ChairNumber']}")
            print(f" Location: {report['Location']}")
            print(f" Description: {report['Description']}")
            print(f" Reported by: {report['Informant']}")
            print(f" Date: {report['Date']}")
            print(f" Status: {report['Status']}")
        print("\nTotal Damaged Items Reported:", len(damage_reports), "\n")

def update_status():
    view_reports()
    if len(damage_reports) == 0:
        return
    
    try:
        report_num = int(input("Enter report number to update status: "))
        if any(r["ReportNumber"] == report_num for r in damage_reports):
            new_status = input("Enter New Status (Fixed / Pending / Follow-up): ").capitalize()
            if new_status in ["Fixed", "Pending", "Follow-up"]:
                for report in damage_reports:
                    if report["ReportNumber"] == report_num:
                        report["Status"] = new_status
                        print(f"\nReport {report_num} status updated to {new_status}.\n")
                        save_report_to_excel(report)
                        break
            else:
                print("Invalid Status. Please Enter 'Fixed', 'Pending', or 'Follow-up'.")
        else:
            print("Invalid Report Number.")
    except ValueError:
        print("Please Enter a Valid Number.")

def summary_report():
    if not os.path.exists(excel_file):
        print("\nNo data found. Please create reports first.\n")
        return

    wb = load_workbook(excel_file)
    sheet = wb.active

    total_reports = 0
    damaged_items = {}
    chair_numbers = []

    # Skip the header row (start from row 2)
    for row in sheet.iter_rows(min_row=2, values_only=True):
        total_reports += 1
        item = row[1]  # "Item" column
        chair_num = row[2]  # "Chair Number" column

        # Count each type of item
        if item in damaged_items:
            damaged_items[item] += 1
        else:
            damaged_items[item] = 1

        # Store chair number if applicable
        if item.lower() == "chair" and chair_num != "N/A":
            chair_numbers.append(chair_num)

    # Print summary
    print("\n=== Summary Report ===")
    print(f"Total Damaged Items Reported: {total_reports}")

    print("\nItems Breakdown:")
    for item, count in damaged_items.items():
        print(f" {item}: {count}")

    if chair_numbers:
        print("\nDamaged Chair Numbers:", ", ".join(chair_numbers))

    print("\n======================\n")

def main():
    while True:
        print("=== Damaged Items Report System ===")
        print("1. Report Damaged Item")
        print("2. View All Reports")
        print("3. Update Report Status")
        print("4. Generate Summary Report")
        print("5. Exit")
        choice = input(f"\nEnter Your Choice: ")

        if choice == "1":
            report_damage()
        elif choice == "2":
            view_reports()
        elif choice == "3":
            update_status()
        elif choice == "4":
            summary_report()
        elif choice == "5":
            break
        else:
            print("Invalid choice, please try again.\n")

if __name__ == "__main__": main()