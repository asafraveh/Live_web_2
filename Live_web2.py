import tkinter as tk
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


def open_website_and_login(website_url, username_value, password_value):
    # Create a new Chrome browser instance
    driver = webdriver.Chrome()

    # Open the specified website
    driver.get(website_url)

    # Use explicit wait to wait for the username field to be clickable
    wait = WebDriverWait(driver, 10)
    username_field = wait.until(EC.element_to_be_clickable((By.ID, "username")))

    # Click the username field to focus it
    username_field.click()

    # Send keys to the username field
    username_field.send_keys(username_value)

    # Similarly, handle the password field and login button

    password_field = wait.until(EC.element_to_be_clickable((By.ID, "password")))
    password_field.click()
    password_field.send_keys(password_value)

    login_button = wait.until(EC.element_to_be_clickable((By.ID, "kc-login")))
    login_button.click()

    # Wait for user confirmation before closing the browser
    input("Press Enter to close the browser...")
    driver.quit()  # Close the browser when the user presses Enter


# Create the main window
root = tk.Tk()
root.title("Website Automation")

# Create buttons
button1 = tk.Button(root, text="CEMEX", command=lambda: open_website_and_login(
    "https://cemex-manager-app.ception.live/",
    "ceptionuser",
    "ceptioncemex?"
))
button2 = tk.Button(root, text="Shafir", command=lambda: open_website_and_login(
    "https://shapir-manager-app.ception.live/",
    "ceptionuser",
    "ceptionshapir!"
))
button3 = tk.Button(root, text="Heidelberg", command=lambda: open_website_and_login(
    "https://heidelberg-manager-app.ception.live/",
    "ceptionuser",
    "ceptionheidelberg%"
))

# Pack buttons into the main window
button1.pack()
button2.pack()
button3.pack()

# Start the GUI main loop
root.mainloop()
