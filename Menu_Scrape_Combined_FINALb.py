import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
import time
import os
from datetime import date, datetime
import pytz

# Dictionary mapping state names to their 2-letter codes
state_codes = {
    "Alabama": "AL", "Alaska": "AK", "Arizona": "AZ", "Arkansas": "AR",
    "California": "CA", "Colorado": "CO", "Connecticut": "CT", "Delaware": "DE",
    "Florida": "FL", "Georgia": "GA", "Hawaii": "HI", "Idaho": "ID",
    "Illinois": "IL", "Indiana": "IN", "Iowa": "IA", "Kansas": "KS",
    "Kentucky": "KY", "Louisiana": "LA", "Maine": "ME", "Maryland": "MD",
    "Massachusetts": "MA", "Michigan": "MI", "Minnesota": "MN", "Mississippi": "MS",
    "Missouri": "MO", "Montana": "MT", "Nebraska": "NE", "Nevada": "NV",
    "New Hampshire": "NH", "New Jersey": "NJ", "New Mexico": "NM", "New York": "NY",
    "North Carolina": "NC", "North Dakota": "ND", "Ohio": "OH", "Oklahoma": "OK",
    "Oregon": "OR", "Pennsylvania": "PA", "Rhode Island": "RI", "South Carolina": "SC",
    "South Dakota": "SD", "Tennessee": "TN", "Texas": "TX", "Utah": "UT",
    "Vermont": "VT", "Virginia": "VA", "Washington": "WA", "West Virginia": "WV",
    "Wisconsin": "WI", "Wyoming": "WY", "Washington DC": "DC"
}

def get_toastique_menu(url):
    try:
        time.sleep(1)  # This is the line you added

        chrome_options = Options()
        chrome_options.add_argument("--headless")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--window-size=1920x1080")

        service = Service()
        driver = webdriver.Chrome(service=service, options=chrome_options)

        driver.delete_all_cookies()
        driver.get(url)
        time.sleep(2)

        page_source = driver.page_source
        soup = BeautifulSoup(page_source, "html.parser")

        restaurant_name_element = soup.find("div", class_="restaurant-name")
        if restaurant_name_element:
            restaurant_name = restaurant_name_element.text.strip()
        else:
            restaurant_name = "Restaurant Name Not Found"

        menu_items_data = []

        menu_items = soup.find_all("div", class_="item-card")

        if not menu_items:
            return restaurant_name, None

        for item in menu_items:
            try:
                name_element = item.find("div", class_="item-card-name")
                price_element = item.find("div", class_="item-card-price")

                if name_element and price_element:
                    name = name_element.text.strip()
                    price = price_element.text.strip()
                    menu_items_data.append({"name": name, "price": price})
            except AttributeError:
                pass

        return restaurant_name, menu_items_data

    except Exception as e:
        return f"Error: {e}", None

    finally:
        if 'driver' in locals() and driver:
            driver.quit()

def save_combined_excel(menu_data_dict, folder_path, start_time, master_menu=None):
    """Saves menu data into separate sheets based on state and highlights price differences."""

    workbook = openpyxl.Workbook()
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    light_blue_fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid") # Light Blue

    for state, menu_data_list in menu_data_dict.items():
        state_code = state_codes.get(state, state)  # Get state code or keep original state name if not found
        sheet_title = state_code.replace(":", "")
        sheet = workbook.create_sheet(title=sheet_title)

        col_index = 1

        for restaurant_name, menu_data in menu_data_list:
            if not menu_data:
                continue

            sheet.merge_cells(start_row=1, start_column=col_index, end_row=1, end_column=col_index + 1)
            restaurant_cell = sheet.cell(row=1, column=col_index, value=restaurant_name)
            restaurant_cell.alignment = Alignment(horizontal='center')
            restaurant_cell.font = Font(bold=True, size=16)

            item_name_header = sheet.cell(row=2, column=col_index, value="Item Name")
            price_header = sheet.cell(row=2, column=col_index + 1, value="Price")
            item_name_header.font = Font(bold=True, size=14)
            price_header.font = Font(bold=True, size=14)

            if menu_data:
                for row_index, item in enumerate(menu_data, start=3):
                    item_name_cell = sheet.cell(row=row_index, column=col_index, value=item["name"])
                    item_price_cell = sheet.cell(row=row_index, column=col_index + 1, value=item["price"])

                    if master_menu:
                        if item["name"] not in master_menu:
                            item_name_cell.fill = light_blue_fill
                            item_price_cell.fill = light_blue_fill
                        else:
                            master_item_price = master_menu.get(item["name"])
                            if master_item_price:
                                if item["price"] != master_item_price:
                                    try:
                                        current_price = float(item["price"])
                                        master_price = float(master_item_price)
                                        if current_price > master_price:
                                            item_name_cell.fill = red_fill
                                            item_price_cell.fill = red_fill
                                        elif current_price < master_price:
                                            item_name_cell.fill = yellow_fill
                                            item_price_cell.fill = yellow_fill
                                    except ValueError:
                                        # Handle cases where price might not be a valid number
                                        print(f"Warning: Could not compare prices for '{item['name']}' at {restaurant_name} due to invalid price format.")

            col_index += 3

        for col in sheet.columns:
            max_length = 0
            column = get_column_letter(col[0].column)
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2) * 1.1
            sheet.column_dimensions[column].width = adjusted_width

    if "Sheet" in workbook.sheetnames and len(workbook["Sheet"]._cells) == 0:
        del workbook["Sheet"]

    dc_timezone = pytz.timezone('America/New_York')
    dc_time = datetime.now(dc_timezone).strftime("%I'%M%p")

    filename = f"Toastique Menu Master File - {dc_time}.xlsx"
    file_path = os.path.join(folder_path, filename)

    try:
        workbook.save(file_path)
        print(f"\nSuccessfully Retrieved {sum(len(v) for v in menu_data_dict.values())} Toastique Menus!")
        print(f"The Menu Master File has been saved to {file_path}")

    except Exception as e:
        print(f"Error saving combined menu data: {e}")

# Main execution
today = date.today().strftime("%m-%d-%Y")
folder_path = today
os.makedirs(folder_path, exist_ok=True)

menu_data_dict = {}
start_time = time.time()
failed_links = []
master_menu = {}
master_store_name = "Toastique Union Market" # Updated master store name

try:
    with open("Menu Links.txt", "r") as file:
        current_state = None
        for line in file:
            line = line.strip()
            if not line or line.startswith("#"):
                continue

            if "http" not in line:
                current_state = line
                menu_data_dict[current_state] = []
            else:
                parts = line.split(" - ")
                if len(parts) == 2:
                    location_name, url = parts
                    restaurant_name, menu_data = get_toastique_menu(url)
                    if menu_data:
                        if restaurant_name == master_store_name:
                            for item in menu_data:
                                master_menu[item["name"]] = item["price"]
                            menu_data_dict[current_state].append((restaurant_name, menu_data)) # Add Clarendon to the main data
                            print(f"Successfully Retrieved the Master Menu for {restaurant_name} ({state_codes.get(current_state, current_state)})!")
                        else:
                            menu_data_dict[current_state].append((restaurant_name, menu_data))
                            print(f"Successfully Retrieved the Menu for Toastique {location_name} ({state_codes.get(current_state, current_state)})!")
                    else:
                        failed_links.append((current_state, location_name, url))
                        print(f"Failed to retrieve menu data for {location_name} ({current_state})")
                else:
                    print(f"Warning: Invalid line format in Menu Links.txt: {line}")

except FileNotFoundError:
    print("Error: Menu Links.txt not found.")
except Exception as e:
    print(f"An unexpected error occurred: {e}")

# Retry failed links
if failed_links:
    print("\nRetrying failed links...")
    for state, location_name, url in failed_links:
        restaurant_name, menu_data = get_toastique_menu(url)
        if menu_data:
            if restaurant_name != master_store_name: # Keep this condition to avoid duplicates if retry succeeds for Clarendon
                menu_data_dict[state].append((restaurant_name, menu_data))
                print(f"Successfully Retrieved the Menu for Toastique {location_name} ({state_codes.get(current_state, current_state)}) (Retry)!")
        else:
            print(f"Failed to retrieve menu data for {location_name} ({state}) (Retry).")

save_combined_excel(menu_data_dict, folder_path, start_time, master_menu)

if master_menu:
    print(f"\nComparison for prices was done against {master_store_name}")
elif menu_data_dict:
    print("\nNo master store menu was retrieved, so no price comparison was done.")
else:
    print("\nNo menu data was retrieved.")

end_time = time.time()
elapsed_time = end_time - start_time
print(f"\nThe script took {elapsed_time:.2f} seconds to complete.")
