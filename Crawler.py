import tkinter as tk
from tkinter import filedialog, messagebox, Label
import requests
from bs4 import BeautifulSoup
import pandas as pd
from threading import Thread


def get_element_text(section, text_label):
    element = section.find(text=lambda t: t and text_label in t)
    if element:
        return element.find_next().text.strip()
    return ""

def scrape_data(file_path, max_empty_pages=100, update_progress=None, finished_callback=None):
    base_url = "https://companyinfo.ge/en/corporations/{}"
    company_id = 1
    empty_page_count = 0
    
    df = pd.DataFrame(columns=["Company ID", "Legal Form", "Registration Date", "Source Information", "Address", "Email", "Status"])
    
    try:
        while empty_page_count < max_empty_pages:
            url = base_url.format(company_id)
            response = requests.get(url)
            
            if response.status_code == 200 and response.text:
                soup = BeautifulSoup(response.text, 'html.parser')
                about_section = soup.find('div', class_='details-section about')


                if about_section:
                    empty_page_count = 0
                    # Use the get_element_text function for safe data extraction
                    company_data = {
                        "Company ID": company_id,
                        "Legal Form": get_element_text(about_section, 'Legal Form:'),
                        "Registration Date": get_element_text(about_section, 'Registration Date:'),
                        "Source Information": get_element_text(about_section, 'Source information:'),
                        "Address": get_element_text(about_section, 'Address:'),
                        "Email": get_element_text(about_section, 'Email:'),
                        "Status": get_element_text(about_section, 'Status:')
                    }
                    df = pd.concat([df, pd.DataFrame([company_data])], ignore_index=True)


                else:
                    empty_page_count += 1
            else:
                empty_page_count += 1
            
            company_id += 1
            if update_progress:
                update_progress(company_id)
        
        df.to_excel(file_path, index=False)
        
        if finished_callback:
            finished_callback(True, f"Data collected and saved to {file_path}")
    except Exception as e:
        if finished_callback:
            finished_callback(False, str(e))

def update_progress(current_id):
    progress_var.set(f"Scraping company ID: {current_id}...")

def finished_scraping(success, message):
    if success:
        messagebox.showinfo("Completed", message)
    else:
        messagebox.showerror("Error", message)
    save_button.config(state=tk.NORMAL)

def start_scraping_thread(file_path):
    save_button.config(state=tk.DISABLED)
    thread = Thread(target=scrape_data, args=(file_path,), kwargs={'update_progress': update_progress, 'finished_callback': finished_scraping})
    thread.start()

def save_file():
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        start_scraping_thread(file_path)

root = tk.Tk()
root.title("Web Scraping Tool")

frame = tk.Frame(root)
frame.pack(padx=10, pady=10)

save_button = tk.Button(frame, text="Save As...", command=save_file)
save_button.pack(side=tk.LEFT)

progress_var = tk.StringVar(root)
progress_label = Label(root, textvariable=progress_var)
progress_label.pack(pady=10)

root.mainloop()
