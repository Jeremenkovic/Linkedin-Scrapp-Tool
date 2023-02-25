from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup
import time
import pandas as pd
import tkinter as tk
from tkinter import messagebox
from tkinter.ttk import Progressbar

def scrape_connections(progress):
    options = webdriver.ChromeOptions()
    options.add_argument('headless')
    driver = webdriver.Chrome(options=options)

    driver.get('https://www.linkedin.com/login')
    username_input = driver.find_element_by_id('username')
    username_input.send_keys('YOUR_USERNAME')
    password_input = driver.find_element_by_id('password')
    password_input.send_keys('YOUR_PASSWORD')
    password_input.send_keys(Keys.RETURN)
    time.sleep(5)

    driver.get('https://www.linkedin.com/mynetwork/invite-connect/connections/')
    time.sleep(5)

    columns = ['Name', 'Title', 'Company', 'Profile Link']
    data = pd.DataFrame(columns=columns)

    while True:
        driver.execute_script('window.scrollTo(0, document.body.scrollHeight);')
        time.sleep(2)

        html = driver.page_source
        soup = BeautifulSoup(html, 'html.parser')

        connections = soup.find_all('li', {'class': 'mn-connection-card'})

        for connection in connections:
            name = connection.find('span', {'class': 'mn-person-info__name'}).text.strip()
            title = connection.find('span', {'class': 'mn-person-info__occupation'}).text.strip()
            company = connection.find('span', {'class': 'mn-person-info__company'}).text.strip()
            profile_link = connection.find('a', {'class': 'mn-connection-card__link'})
            if profile_link:
                profile_link = 'https://www.linkedin.com' + profile_link['href']
            else:
                profile_link = None

            if 'general manager' in title.lower() or 'ceo' in title.lower() or 'software' in company.lower():
                if profile_link:
                    data = data.append({'Name': name, 'Title': title, 'Company': company, 'Profile Link': profile_link}, ignore_index=True)
                print(f'{name} ({title}, {company}): {profile_link}')

        load_more_button = driver.find_element_by_xpath('//button[text()="Show more"]')
        if not load_more_button.is_enabled():
            break
        load_more_button.click()
        time.sleep(2)
        progress['value'] += 10
        window.update_idletasks()

    data.to_excel('connections.xlsx', index=False)

    driver.quit()

    progress['value'] = 100
    window.update_idletasks()

    messagebox.showinfo('Scraping Complete', 'The scraping is complete and the data has been saved to "connections.xlsx".')

def on_scrape_click():
    scrape_button.config(state='disabled')
    progress_bar['value'] = 0
    progress_bar['maximum'] = 100

    scrape_connections(progress_bar)

    scrape_button.config(state='normal')
    progress_bar['value'] = 0

window = tk.Tk()
window.title('LinkedIn Scraper')
window.geometry('300x150')

scrape_button = tk.Button(window, text='Scrape Connections', command=on_scrape_click)
scrape_button.pack(pady=10)

progress_bar = Progressbar(window, orient=tk.HORIZONTAL, length=200, mode='determinate')
progress_bar.pack(pady=10)

window.mainloop()
