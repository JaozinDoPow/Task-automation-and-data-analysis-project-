# Library
import time 
import pyautogui as pag # Automation
import pandas # Data Analysis

# Variables
user = "user" # Name of the person who will send the email to the company.
database_link = "https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga" # A Link that simulates the link to the company's database.
archive_path = r"C:\Users\user\Downloads\Vendas - Dez.xlsx" # The archive path.
max_sales = 35450 # In my case, that's the max sales amount
min_sales = 0

# Opening Browser
pag.PAUSE = 0.4 
pag.press("win") 
pag.write("Chrome") # The Browser, might be a different one.
pag.press("enter") 
time.sleep(0.75) 

#  Acessing System
pag.write(database_link)
pag.press("enter")
time.sleep(2) 

# Searching and Downloanding the archive
pag.click(x = 400, y = 450, clicks = 2) # The pixels that need to be clicked (on my screen)
time.sleep(0.7) 
pag.click(x = 400, y = 450, button = "right") 
pag.click(x = 460, y = 510)
time.sleep(3)
pag.press("enter")
time.sleep(1)

# Acessing data in the archive
spreadsheet = pandas.read_excel(archive_path) # The Archive is xlsx.
total_sales = spreadsheet["Quantidade"].sum() # "Quantidade" = Amount
revenue = spreadsheet["Valor Final"].sum() # "Valor Final" = Final Value (Price)

def asses_sales(sales): # A function to asses the sales
    max = 35450 
    min = 0 
    performance = ""

    percentage = (sales - min) / (max - min) 
    if percentage < 0.25: # Use one word to describe how good the sales were.
        performance = "bad"
    elif percentage < 0.5:
        performance = "neutral"
    elif percentage < 0.75:
        performance = "good"
    else:
        performance = "excellent"
    return percentage * 100, performance  # Return in percentage and in a word

percentage, performance = asses_sales(total_sales)

# Send email with the report and the data
title = f"Sales Report - {user}"
pag.hotkey("ctrl", "t") 
pag.write("https://mail.google.com/") # Using gmail, in my case
pag.press("enter")
time.sleep(2)
pag.click(x = 110, y = 220) 
time.sleep(1)
pag.write("joaolmsantana+relatorios@gmail.com") # Writes the recipient email, my email for example
pag.press("tab", presses = 2) 
pag.write(title) 
pag.press("tab") 
text = f"""Dear Sir/Madam,

Here is today's sales report.

Total Sales: {total_sales}.

Revenue: R$ {revenue:.2f}.

We sold {percentage:.2f}% of the total products we could sell, and we had a {performance} number of sales.

Please feel free to contact me with any questions.

Sincerely, {user}"""

pag.write(text) 
pag.hotkey("ctrl", "enter") 