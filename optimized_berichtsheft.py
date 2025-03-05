import datetime
import os
from pathlib import Path
from PyPDF2 import PdfReader
import openpyxl
import requests
import sys
from bs4 import BeautifulSoup
import subprocess


class Berichtsheftmaker:
    def __init__(self):
        self.calenderweek = datetime.datetime.now().isocalendar()[1]
        self.currentyear = datetime.date.today().year
        self.output_folder = f"KW{self.calenderweek}"
        self.repo_path = os.path.dirname(os.path.abspath(__file__))
        

        if os.path.exists(self.output_folder):
            print(f"Folder {self.output_folder} already exists. Exiting.")
            sys.exit(0)
            
        os.makedirs(self.output_folder, exist_ok=True)
        self.download_mass_pdf()
    
    def get_pdf_url(self):
        website = "https://service.viona24.com/stpusnl"
        response = requests.get(website)
        soup = BeautifulSoup(response.text, "html.parser")
        thelist = soup.find(id="thelist")
        thelist_items = thelist.find_all("a")
        
        # More efficient dictionary comprehension
        return {
            item.text.strip("\n"): item.get("href").strip("./")
            for item in thelist_items
        }
    
    def download_mass_pdf(self):
        listdict = self.get_pdf_url()
        for key, value in listdict.items():
            if "US" in key:
                url = f"https://service.viona24.com/stpusnl/{value}"
                filename = f"{key}.pdf"
                
                # Download file
                response = requests.get(url)
                Path(filename).write_bytes(response.content)
                print(f"Downloaded {filename}")
                
                # Process file
                self.process_pdf(filename, key)
            
    def process_pdf(self, file, key):
        """Combined PDF processing to reduce file operations"""
        # Extract text from PDF
        reader = PdfReader(file)
        text = ""
        for page in reader.pages:
            text += page.extract_text()
            
        # Process text directly without writing to intermediate file
        unedited_list = []
        for line in text.splitlines():
            if "Teams" in line or "Extern" in line:
                continue
            elif " / " in line and ":" not in line:
                fach = line.strip("\n").strip(",")
                unedited_list.append(fach)
            elif "15:00-16:00" in line or "16:15" in line:
                unedited_list.append(".")
            elif "Mentor" in line and "Verf" in line:
                unedited_list.append("Verfügungsstd.")
        
        # Remove duplicates
        edited_list = []
        for i, fach in enumerate(unedited_list):
            if i == 0 or fach != unedited_list[i-1]:
                edited_list.append(fach)
        
        # Create Excel file
        self.listtoexcel(edited_list, f"{file}-output.txt")
        
        # Clean up PDF file
        os.remove(file)

    def git_commit_and_push_folder(self):
        try:
            os.chdir(self.repo_path)
            subprocess.run(['git', 'add', self.output_folder], check=True)
            commit_message = f"Automated commit: Added {self.output_folder}"
            subprocess.run(['git', 'commit', '-m', commit_message], check=True)
            subprocess.run(['git', 'push'], check=True)
            
            print(f"Folder {self.output_folder} committed and pushed to the repository.")
        except subprocess.CalledProcessError as e:
            print(f"Git operation failed: {e}")

    def listtoexcel(self, data, filename):
        name = filename.removesuffix(".pdf-output.txt")
        
        # Get base path for template file
        if getattr(sys, 'frozen', False):
            base_path = sys._MEIPASS
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))
            
        path = os.path.join(base_path, "copycopy.xlsx")
        
        # Load workbook and prepare variables
        workbook = openpyxl.load_workbook(path)
        worksheet = workbook["Tabelle1"]
        counter = 4
        daycounter = 0
        praxiszeit = 420
        std_dauer = 90
        pausen_counter = 9
        praxis_counter = 8
        
        try:
            for fach in data:
                if "Ver" not in fach and "." not in fach:
                    worksheet[f"B{counter}"] = f"{fach}:"
                    worksheet[f"E{counter}"] = std_dauer
                    praxiszeit -= std_dauer
                    counter += 1
                elif "Ver" in fach:
                    worksheet[f"B{counter}"] = f"{fach}:"
                    worksheet[f"E{counter}"] = std_dauer / 2
                    praxiszeit -= std_dauer / 2
                    counter += 1
                if "." in fach and "Ver" not in fach:
                    counter = 10 + (6 * daycounter)
                    daycounter += 1
                    worksheet[f"B{praxis_counter}"] = "Praxisunterricht:"
                    worksheet[f"E{praxis_counter}"] = praxiszeit
                    worksheet[f"E{pausen_counter}"] = 60
                    praxis_counter += 6
                    pausen_counter += 6
                    praxiszeit = 420
        except AttributeError:
            print(f"Error at {counter}")
            
        # Set header information
        worksheet["D1"] = f"KW{self.calenderweek}"
        worksheet["E1"] = f"Jahr {self.currentyear}"
        
        # Save workbook
        output_path = os.path.join(self.output_folder, f"{name}-Berichtsheft_KW{self.calenderweek}.xlsx")
        workbook.save(output_path)
        print(f"Excel file saved to: {output_path}")

    def __del__(self):
        # Commit and push the folder when the program ends
        self.git_commit_and_push_folder()


if __name__ == "__main__":
    Berichtsheftmaker() 