import re
import pandas as pd
from openpyxl import load_workbook
from pypdf import PdfReader
import logging

logging.getLogger("pypdf").setLevel(logging.ERROR)  # Suppress warnings

def validate_pages(pages):
    """Validate if the user input follows the page range format."""
    pages = pages.replace(" ", "")
    pattern = r"^\d+(-\d+)?(,\d+(-\d+)?)*$"  # Supports single pages and ranges
    return bool(re.match(pattern, pages))

def parse_pages(pages):
    """Parse page input and return a sorted list of page numbers."""
    page_set = set()
    pages = pages.replace(" ", "")

    for part in pages.split(","):
        if "-" in part:
            start, end = map(int, part.split("-"))
            page_set.update(range(start, end + 1))
        else:
            page_set.add(int(part))

    return sorted(page_set)

def readPDFPages(page_list, file_path=""): # Fill with file path
    """Extract text from the specified pages in the PDF."""
    reader = PdfReader(file_path)
    extracted_text = ""

    for page_num in page_list:
        try:
            page = reader.pages[page_num - 1]  # Convert to zero-based index
            text = page.extract_text()
            if text:
                extracted_text += text + "\n"
        except IndexError:
            print(f"Warning: Page {page_num} is out of range.")

    return extracted_text.strip() if extracted_text.strip() else None

def parse_roster(text):
    EXCLUDE_WORDS = {"State", "Hometown", "Ecol", "Speech", "Language", "Barnechea, Santiago", "Evol", "Environ Biol", "Republic", "Korea","Mon", "Tues", "Wed", "Thurs", "Fri", "Sat", "Sun", "Feb", "Sep", "MCUT-N2", "MCUT-C2", "MCUT-N1", "MCUT-N2", "MCUT-N3", "MCUT-N4", "MCUT-N5", "MCUT-N6", "MCUT-N7", "MCUT-N8", "MCUT-S1", "MCUT-S2", "MCUT-S3", "MCUT-S4", "MCUT-S5", "MCUT-S6", "MCUT-S7", "MCUT-S8", "MCUT-C1"}
    
    # Pattern for matching names
    name_pattern = r"([A-Z][a-z]+, [A-Z][a-z]+(?: [A-Z])?)"

    # Pattern for matching email addresses
    email_pattern = r"(\S+?@\S+?\.\S+)"
    
    # Pattern for matching room numbers (e.g., MCUT-123)
    room_pattern = r"(MCUT-\S+)"
    
    names_and_emails_and_rooms = []
    
    # Find all names
    names = re.findall(name_pattern, text)

    # Find all email addresses
    emails = re.findall(email_pattern, text)

    # Find all room numbers
    rooms = re.findall(room_pattern, text)

    lisOfFirstNames = []
    lisOfLastNames = []
    lisOfEmails = []
    lisOfRoomss = []

    for i in range(len(names)):
        if names[i] not in EXCLUDE_WORDS:
            last, first = names[i].split(", ")
            first_name = first.strip()
            last_name = last.strip()
            if first_name not in EXCLUDE_WORDS:
                lisOfFirstNames.append(first_name)
            if last_name not in EXCLUDE_WORDS:
                lisOfLastNames.append(last_name)

    for i in range(len(emails)):
        if emails[i] not in EXCLUDE_WORDS:
            lisOfEmails.append(emails[i])

    for i in range(len(rooms)):
        if rooms[i] not in EXCLUDE_WORDS:
            lisOfRoomss.append(rooms[i])

    print(len(lisOfFirstNames), len(lisOfLastNames), len(lisOfEmails), len(lisOfRoomss))
    print(lisOfFirstNames)
    print()
    print(lisOfLastNames)
    print()
    print(lisOfEmails)
    print()
    print(lisOfRoomss)
    return lisOfFirstNames, lisOfLastNames, lisOfEmails, lisOfRoomss


def moveToExcel(first, last, email, room, filename="roster.xlsx"):
    """Write extracted roster data to an Excel file and adjust column width."""
    
    # Create a DataFrame
    df = pd.DataFrame({
        "First Name": first,
        "Last Name": last,
        "Email": email,
        "Room Number": room
    })
    
    # Write DataFrame to an Excel file with openpyxl engine
    with pd.ExcelWriter(filename, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Roster")

        # Get the workbook and worksheet
        workbook = writer.book
        worksheet = writer.sheets["Roster"]
        
        # Set column widths
        column_widths = {
            "A": 15,  # First Name
            "B": 15,  # Last Name
            "C": 30,  # Email
            "D": 12   # Room Number
        }

        for col, width in column_widths.items():
            worksheet.column_dimensions[col].width = width

def main():
    """Main function to run the program."""
    while True:
        pages = input("Enter the pages you want to import (e.g., '1-3' or '1, 3, 5'): ").strip()

        if validate_pages(pages):
            page_list = parse_pages(pages)
            break
        else:
            print("Invalid input. Please enter a valid range (e.g., '1-3') or a comma-separated list (e.g., '1, 3, 5').")

    text = readPDFPages(page_list)

    if not text:
        print("Error: No text could be extracted from the selected pages.")
        return

    first,last,email,room = parse_roster(text)
    moveToExcel(first,last,email,room)

if __name__ == "__main__":
    main()
