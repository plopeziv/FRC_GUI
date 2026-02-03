# -*- coding: utf-8 -*-
"""
Created on Fri Dec 12 10:11:42 2025

@author: plopez
"""

import re
import os
import win32com.client as win32
from datetime import datetime
from pathlib import Path
from PyPDF2 import PdfWriter

from utils.date_utils import parse_string_date

def process_ticket(input_path, date, ticket_status):
    pdf_location = excel_to_pdf(input_path)
    
    job_folder = find_job_folder(input_path)
    
    job_number = job_folder.name
    ticket_number = extract_ticket_number(input_path)
    
    job_ticket = find_ticket_file(job_folder,  ticket_number)
    
    bid_folder = find_bid_folder(job_number) / "TICKETS"
    bid_folder.mkdir(exist_ok=True)
    
    formatted_date = parse_string_date(date.replace("/","-")).strftime("%m-%d-%y")
    
    display_status = "SIGNED" if ticket_status.upper() == "YES" else "UNSIGNED"

    final_pdf_path = str(bid_folder / f"{job_number} {formatted_date} {ticket_number} {display_status}.pdf")
     
    merge_pdfs(
        str(pdf_location),
        str(job_ticket),
        final_pdf_path
        )

    print(f"Merged PDF saved at {final_pdf_path}")

    os.startfile(final_pdf_path)

def excel_to_pdf(input_path):
    input_path = Path(input_path).resolve()

    pdf_folder = input_path.parent/"pdfs"
    pdf_folder.mkdir(exist_ok=True)
    
    print(f"saving in folder path {pdf_folder}")
    
    output_pdf = pdf_folder / (input_path.stem + ".pdf")

    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    print("Opening Excel Workbook")
    wb = excel.Workbooks.Open(str(input_path))

    # Export as PDF
    print("Saving PDF")
    sheet = wb.Worksheets(1)
    sheet.ExportAsFixedFormat(
        Type=0,                # 0 = PDF
        Filename=str(output_pdf),
        From = 1,
        To = 1
    )

    wb.Close()
    excel.Quit()
    
    print(f"PDF Saved at {output_pdf}")
    
    return output_pdf
    
def merge_pdfs(top_pdf, bottom_pdf, output_path):    
    writer = PdfWriter()

    writer.append(top_pdf)
    writer.append(bottom_pdf)

    with open(output_path, "wb") as f:
        writer.write(f)
        
def extract_ticket_number(path):
    name = Path(path).stem

    match = re.search(r"-\s*(\d+)", name)
    if not match:
        raise ValueError(f"Could not extract ticket number from '{name}'")

    return match.group(1)
    
def find_job_folder(input_path):
    WORK_STATION = Path(r"F:\USERS\Pedro L\Ticket Work Station")
    input_path = Path(input_path)
    
    job_number = input_path.stem.split("-", 1)[0].strip()
    
    folder_path = WORK_STATION / job_number
    
    
    if not folder_path.exists() or not folder_path.is_dir():    
        raise FileNotFoundError(
            f"{job_number} not found in Work Station folder"
        )
        
    return folder_path

def find_bid_folder(job_number):
    BIDS_FOLDER = Path(r"\\FRC2\otherapps\Doc_Arch\Project Folders\0 Structure\Bids")
    
    # Try with space after dash first
    matches = list(BIDS_FOLDER.glob(f"{job_number} -*"))
    
    # If nothing, try without space
    if not matches:
        matches = list(BIDS_FOLDER.glob(f"{job_number}-*"))
    
    if not matches:
        raise FileNotFoundError(
            f"'{job_number}' not found in {BIDS_FOLDER}"
        )
        
    # Return the first matching folder
    return matches[0]


def find_ticket_file(job_folder, ticket_number):
    job_folder = Path(job_folder)
    
    patterns = [
        f"TICKET {ticket_number} SIGNED*.pdf",
        f"TICKET {ticket_number} UNSIGNED*.pdf"
    ]
    
    for pattern in patterns:
        matches = list(job_folder.glob(pattern))
        if matches:
            return matches[0]
        
    raise FileNotFoundError(
        f"No pdf found for Ticket {ticket_number}"
        )

if __name__ == "__main__":
    input_file = r"C:\Users\plopez\Desktop\FRC_GUI\data_manager\027386 - 040691.xlsx"
    process_ticket(input_file, "4/06/91", "UNSIGNED")
