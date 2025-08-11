import os
import requests
import undetected_chromedriver as uc
from bs4 import BeautifulSoup
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from urllib.parse import urljoin
from PIL import Image as PILImage
from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage
import json
import random

# Function to resize image with quality preservation
def resize_image(image_path, max_width=500, max_height=500):
    try:
        with PILImage.open(image_path) as img:
            img.thumbnail((max_width, max_height), PILImage.ANTIALIAS)
            img.save(image_path, format="JPEG", quality=95)
            print(f"Resized and saved image: {image_path}")
    except Exception as e:
        print(f"Error resizing image {image_path}: {e}")

# Function to download image and save it locally in a product-specific folder under 'public' folder
def download_image(image_url, product_folder, image_name):
    try:
        # Ensure the URL is complete
        if not image_url.startswith('http'):
            image_url = urljoin("https://atris.com.au", image_url)

        response = requests.get(image_url, stream=True)
        response.raise_for_status()

        os.makedirs(product_folder, exist_ok=True)  # Ensure product-specific folder exists

        image_path = os.path.join(product_folder, image_name)

        with open(image_path, 'wb') as file:
            for chunk in response.iter_content(1024):
                file.write(chunk)

        print(f"Image saved: {image_path}")
        resize_image(image_path)  # Resize image after downloading
        return image_path
    except Exception as e:
        print(f"Error downloading image {image_url}: {e}")
        return None

# Function to read already scraped links
def read_scraped_links(filename):
    if os.path.exists(filename):
        with open(filename, 'r') as file:
            return set(line.strip() for line in file)
    return set()

# Function to save the scraped link
def save_scraped_link(filename, link):
    with open(filename, 'a') as file:
        file.write(link + '\n')

# Function to store scraped data in JSON
def save_data_to_json(data, json_file):
    try:
        with open(json_file, 'w') as jsonf:
            json.dump(data, jsonf, indent=4)
        print(f"Data saved to {json_file}")
    except Exception as e:
        print(f"Error saving data to JSON: {e}")


def scrape_links_and_save_to_excel(category_name):
    print("Initializing Chrome driver...")
    options = uc.ChromeOptions()
    options.add_argument("--start-maximized")
    driver = uc.Chrome(options=options)

    scraped_links_file = f"scraped_links_{category_name}.txt"
    scraped_links = read_scraped_links(scraped_links_file)

    # Initialize data list to store data for JSON
    scraped_data = []
    base_image_folder = os.path.join("public", "downloaded_images", category_name)  # Save images in public folder
    os.makedirs(base_image_folder, exist_ok=True)

    try:
        # Initialize Excel workbook and sheet
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Scraped Data"
        sheet.append(["ID", "Title", "Category", "SrcUrl", "Gallery", "Rating", "Link"])  # Header row

        input("Navigate to the desired page in the browser and press Enter to start scraping...")

        id_counter = 1
        while True:
            current_url = driver.current_url
            print(f"Scraping page: {current_url}")
            driver.get(current_url)

            # Wait for page to load
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))

            soup = BeautifulSoup(driver.page_source, 'html.parser')
            container_class = "grid grid--uniform grid--scattered-large-4 grid--scattered-small-1"
            container = soup.find('div', class_=container_class)

            if container:
                print("Extracting links...")
                links = [a['href'] for a in container.find_all('a', href=True)]

                for link in links:
                    full_link = f"https://atris.com.au{link}"
                    if full_link in scraped_links:
                        print(f"Skipping already scraped link: {full_link}")
                        continue

                    print(f"Opening link: {full_link}")
                    driver.get(full_link)
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))

                    product_soup = BeautifulSoup(driver.page_source, 'html.parser')

                    # Extract details
                    title = product_soup.find('h1', class_='product-single__title')
                    title = title.get_text(strip=True) if title else "Title not found"
                    product_folder = os.path.join(base_image_folder, title.replace("/", "-"))  # Handle slashes

                    description_wrapper = product_soup.find('div', class_='collapsible-content__inner rte')
                    description = description_wrapper.get_text(separator='\n', strip=True) if description_wrapper else "Description not found"

                    image_container = product_soup.find('div', class_='product__main-photos aos-init aos-animate')
                    gallery = []
                    src_url = None
                    if image_container:
                        image_tags = image_container.find_all('img', src=True)
                        for index, img_tag in enumerate(image_tags):
                            image_url = img_tag['src']
                            image_name = f"image_{index + 1}.jpg"
                            image_path = download_image(image_url, product_folder, image_name)

                            if image_path:
                                if index == 0:  # First image as srcUrl
                                    src_url = image_path
                                gallery.append(image_path)

                    # Assign a random rating between 3.0 and 5.0
                    rating = round(random.uniform(3.0, 5.0), 1)

                    # Prepare data row and JSON object
                    product_data = {
                        "id": id_counter,
                        "title": title,
                        "category": category_name,
                        "description":description,
                        "srcUrl": f"/downloaded_images/{category_name}/{title.replace('/', '-')}/image_1.jpg",  # Relative path
                        "gallery": [f"/downloaded_images/{category_name}/{title.replace('/', '-')}/image_{i+1}.jpg" for i in range(len(gallery))],
                        "rating": rating,
                        "link": full_link,
                    }
                    scraped_data.append(product_data)

                    # Add data to Excel
                    sheet.append([id_counter, title, category_name, product_data['srcUrl'], ", ".join(product_data['gallery']), rating, full_link])

                    save_scraped_link(scraped_links_file, full_link)
                    id_counter += 1
                    print(f"Saved: {full_link}")

            else:
                print("No more content to scrape. Exiting.")
                break

            print("Navigate to the next page in the browser and press Enter...")
            input()

    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        driver.quit()
        excel_file = f"{category_name}_scraped_data.xlsx"
        workbook.save(excel_file)
        print(f"Data saved to {excel_file}")

        # Save data to JSON
        json_file = f"{category_name}_scraped_data.json"
        save_data_to_json(scraped_data, json_file)
        print(f"Data saved to {json_file}")


if __name__ == "__main__":
    category = input("Enter category name (e.g., Surgical Instruments): ").strip()
    scrape_links_and_save_to_excel(category)
