import ctypes
import datetime
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
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
        self.stundenplan = f"StundenplanKW{self.calenderweek}"
        self.output_folder = f"KW{self.calenderweek}"
        self.repo_path = os.path.dirname(os.path.abspath(__file__))
        if os.path.exists(self.output_folder):
            print(f"Folder {self.output_folder} already exists. Exiting.")
            sys.exit(0)
        os.makedirs(self.output_folder, exist_ok=True)
        #pdf: str = self.download_pdf()
        #texfile: str = self.pdf_to_text(pdf)
        #unedited_subjects: list = self.txt_to_list(texfile)
        #cleaned_subjects: list = self.delete_dupe(unedited_subjects)
        #self.listtoexcel(cleaned_subjects)
        #self.send_mail()
        self.download_mass_pdf()
    
    def get_pdf_url(self):
        website = "https://service.viona24.com/stpusnl"
        response = requests.get(website)  # Fetch the HTML content
        soup = BeautifulSoup(response.text, "html.parser")  # Parse the HTML
        thelist = soup.find(id="thelist")
        thelist_items = thelist.find_all("a")
        listdict = {
            item.text.strip("\n"): item.get("href").strip("./")
            for item in thelist_items
        }
        return listdict
    

    def download_mass_pdf(self):
        listdict = self.get_pdf_url()
        for key, value in listdict.items():
            if "US" in key:
                url: str = f"https://service.viona24.com/stpusnl/{value}"
                filename: str = f"{key}.pdf"
                filepath: Path = Path(filename)
                response = requests.get(url)
                filepath.write_bytes(response.content)
                print(f"Downloaded {filename}")
                self.pdf_to_text_mass(filename, key)
                
            else:
                continue

    def pdf_to_text_mass(self, file: str, key: str):
        reader = PdfReader(file)
        text = ""
        for page in reader.pages:
            text += page.extract_text()
        with open(f"{file}-output.txt", "w") as out:
            out.write(text)
        self.txt_to_list_mass(f"{file}-output.txt")
        os.remove(file)

    def txt_to_list_mass(self, file: str):
        with open(file, "r") as file_write:
            unedited_list: list = []
            for line in file_write:
                if "Teams" in line or "Extern" in line:
                    continue
                elif " / " in line and ":" not in line:
                    fach = line.strip("\n").strip(",")
                    unedited_list.append(fach)
                elif "15:00-16:00" in line or "16:15" in line:
                    unedited_list.append(".")
                elif "Mentor" in line and "Verf" in line:
                    unedited_list.append("Verfügungsstd.")
        self.delete_dupe_mass(unedited_list, file)
        os.remove(file)
    
    def delete_dupe_mass(self, data: list, file: str):
        filename = file
        edited_list: list = []
        for fach in range(len(data)):
            if data[fach] != data[fach - 1]:
                edited_list.append(data[fach])
        self.listtoexcel(edited_list, filename)
        
        


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

    def listtoexcel(self, data, filename) -> None:
        filename = filename
        name = filename.removesuffix(".pdf-output.txt")
        if getattr(sys, 'frozen', False):
            base_path = sys._MEIPASS
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))
            
        path = os.path.join(base_path, "copycopy.xlsx")
        
        workbook = openpyxl.load_workbook(path)
        worksheet = workbook["Tabelle1"]
        counter: int = 4
        daycounter: int = 0
        praxiszeit: int = 420
        std_dauer: int = 90
        pausen_counter: int = 9
        praxis_counter: int = 8
        try:
            for fach in data:
                if "Ver" not in fach and "." not in fach:
                    worksheet["B" + str(counter)] = str(fach) + ":"
                    worksheet["E" + str(counter)] = std_dauer
                    praxiszeit -= std_dauer
                    counter += 1
                elif "Ver" in fach:
                    worksheet["B" + str(counter)] = str(fach) + ":"
                    worksheet["E" + str(counter)] = std_dauer / 2
                    praxiszeit -= std_dauer / 2
                    counter += 1
                if "." in fach and "Ver" not in fach:
                    counter = 10 + (6 * daycounter)
                    daycounter += 1
                    worksheet["B" + str(praxis_counter)] = "Praxisunterricht:"
                    worksheet["E" + str(praxis_counter)] = praxiszeit
                    worksheet["E" + str(pausen_counter)] = 60
                    praxis_counter += 6
                    pausen_counter += 6
                    praxiszeit = 420
        except AttributeError:
            print(f"Error at {counter}")
        worksheet["D1"] = f"KW{self.calenderweek}"
        worksheet["E1"] = f"Jahr {self.currentyear}"
        output_path = os.path.join(self.output_folder, f"{name}-Berichtsheft_KW{self.calenderweek}.xlsx")
        workbook.save(output_path)
        print(f"Excel file saved to: {output_path}")

    def send_mail(self) -> None:
        sender_email = os.getenv("SENDER_MAIL")
        password = os.getenv("GMAIL_PASSWORD")
        subject = "mail"
        body = "Body"
        recipient_email = os.getenv("RECIPIENT_MAIL")
        with open(f"Berichtsheft_KW{self.calenderweek}.xlsx", "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition",
                        f"attachment; filename= Berichtsheft_KW{self.calenderweek}.xlsx")
        message = MIMEMultipart()
        message['Subject'] = subject
        message['From'] = sender_email
        message['To'] = recipient_email
        html_part = MIMEText(body)
        message.attach(part)
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(sender_email, password)
            server.sendmail(sender_email, recipient_email, message.as_string())
        os.remove("Berichtsheft_KW47.xlsx")
        print("done")

    def __del__(self):
        # Commit and push the folder when the program ends
        self.git_commit_and_push_folder()

App = Berichtsheftmaker

App()
'''   def download_pdf(self):
        url: str = f"https://service.viona24.com/stpusnl/daten/US_IT_2024_Winter_FIAE_B_2025_abKW{self.calenderweek}.pdf"
        filename: str = f"StundenplanKW{self.calenderweek}.pdf"
        filepath: Path = Path(filename)
        response = requests.get(url)
        filepath.write_bytes(response.content)
        print("success")
        return filename

    @staticmethod
    def pdf_to_text(file: str) -> str:
        reader = PdfReader(file)
        text = ""
        for page in reader.pages:
            text += page.extract_text()
        with open("output.txt", "w") as out:
            out.write(text)
        return "output.txt"

    @staticmethod
    def txt_to_list(data: str) -> list:
        with open(data, "r") as file:
            unedited_list: list = []
            for line in file:
                if "-" in line and ":" not in line:
                    fach = line.split(" ")[0]
                    unedited_list.append(fach)
                elif "16:00" in line:
                    unedited_list.append(".")
                elif "Mentor" in line and "Verf" in line:
                    unedited_list.append("Verfügungsstd.")
        return unedited_list

    @staticmethod
    def delete_dupe(data) -> list:
        edited_list: list = []
        for fach in range(len(data)):
            if data[fach] != data[fach - 1]:
                edited_list.append(data[fach])
            elif data[fach] == ".":
                edited_list.append(".")
        return edited_list'''

