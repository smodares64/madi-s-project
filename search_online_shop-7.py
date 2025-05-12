import aiohttp
import asyncio
from bs4 import BeautifulSoup
import time
import xlwt
from xlwt import Workbook
import xlrd
from xlutils.copy import copy
import os


class Tag:
    """
    A class used to represent a Tag.
    Attributes
    ----------
    data : str
        The data associated with the tag.
    Methods
    -------
    __repr__():
        Returns a string representation of the Tag object.
    """

    def __init__(self, data):
        self.data = data

    def __repr__(self):
        # Optimize this to avoid slow computations
        return f"Tag(data={self.data[:1000000]}...)"


def timer(func):
    """
    A decorator that measures the execution time of a function.
    Args:
        func (callable): The function to be decorated.
    Returns:
        callable: The wrapped function that includes timing functionality.
    Example:
        @timer
        def my_function():
            # Function implementation
            pass
    """

    def wrapper(*args, **kwargs):
        start_time = time.time()
        result = func(*args, **kwargs)
        end_time = time.time()
        print(f"{func.__name__} executed in {end_time - start_time:.2f} seconds")
        return result

    return wrapper


# Function to fetch a page asynchronously
async def get_page(url, retries=8, backoff_factor=2):
    """
    Fetch the HTML page with the categories and products.
    """
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3"
    }

    async with aiohttp.ClientSession() as session:
        for attempt in range(retries):
            try:
                async with session.get(
                    url, headers=headers, timeout=aiohttp.ClientTimeout(total=60000)
                ) as response:
                    response.raise_for_status()  # Raise an error for bad responses (4xx, 5xx)
                    html = await response.text()
                    return BeautifulSoup(html, "html.parser")
            except (
                aiohttp.ClientError,
                aiohttp.ClientResponseError,
                asyncio.TimeoutError,
            ) as e:
                if (
                    attempt == retries - 1
                ):  # If this is the last attempt, raise the exception
                    print(f"Failed to fetch {url} after {retries} attempts. Error: {e}")
                    f = open("demofile2.txt", "a")
                    f.write(url + "\n")
                    f.close()

                    return None
                else:
                    wait_time = backoff_factor * (2**attempt)  # Exponential backoff
                    print(
                        f"Attempt {attempt + 1} failed. Retrying in {wait_time} seconds..."
                    )
                    await asyncio.sleep(wait_time)


# Function to get the sub-category UL tag
def get_sub_category_ul(category: BeautifulSoup):
    """
    Get the UL tag containing sub-categories.
    """
    return category.find("ul", recursive=False)


# Function to get the sub-categories LI tags
def get_sub_categories_li(sub_category_ul: BeautifulSoup):
    """
    Get all LI tags within the sub-category UL.
    """
    return sub_category_ul.find_all("li", recursive=False)


# Function to show sub-categories
def show_sub_categories(category):
    """
    Extract and return the text of the sub-category link.
    """
    return category.select_one("a").text.strip()


# Function to process a URL and extract product details
async def get_stock(link_stock: tuple, index: int, link_stock_list: list):
    """
    Fetch the stock information for a given product link.
    """
    quantity = ""
    if link_stock[0] != "":
        soup = await get_page(link_stock[0])
        if soup:
            quantityAvailable = soup.find("span", id="quantityAvailable")
            if quantityAvailable:
                quantity = quantityAvailable.getText()
            if quantity != "":
                link_stock_list[index] = (link_stock[0], quantity)
                print(
                    f"product in link < {link_stock[0]} > is updated and quantity is {quantity}"
                )


def update_stock_in_excel(
    excelfile: str, link_stock_list: list[tuple], start_row: int = 0, end_row=None
):
    """
    Update stock information in Excel file within specified row range.

    Args:
        excelfile: Path to Excel file
        link_stock_list: List of (link, stock) tuples
        start_row: First row to update (0-based)
        end_row: Last row to update (None for all rows)
    """
    rb = xlrd.open_workbook(excelfile)
    # Create a writable copy
    wb = copy(rb)
    # Get writable sheet
    ws = wb.get_sheet(0)

    if end_row is None:
        # Use full list length if not specified
        end_row = len(link_stock_list)

    i = 0
    for index in range(start_row, end_row):

        if link_stock_list[i][1] == -1:
            # Update the stock in the Excel file
            ws.write(index, 3, link_stock_list[i][1])
            i += 1

    wb.save(excelfile)


async def update_stock(excelfile, start_row: int = 0, end_row=None):
    link_stock_list = await go_to_stock_page(read_excel(excelfile, start_row, end_row))
    update_stock_in_excel(excelfile, link_stock_list, start_row, end_row)


async def run_requests(link_stock_list: list):
    """
    Run asynchronous requests to fetch stock information.
    """
    for index, link_stock in enumerate(link_stock_list):
        await get_stock(link_stock, index, link_stock_list)

    return link_stock_list


# Function to show products in a sub-category
async def get_product_links(href: str, category_name: str, category_list: list[dict]):
    """
    Fetch product links from a given category URL.
    """
    soup = await get_page(href)
    if soup:
        # if there is a form tag, get the url of the button "نمایش همه"
        # ========================================
        link = href + "?"
        form = soup.find("form", class_="showall")
        if form:
            input = form.find(attrs={"name": "id_category"})
            if input:
                link += input["name"] + "=" + input["value"] + "&"
                input2 = form.find(attrs={"name": "n"})
                link += input2["name"] + "=" + input2["value"]
                soup = await get_page(link)
                # ===========================================
                if soup:
                    product_list_ul = soup.find("ul", id="product_list_cat")
                    if product_list_ul:
                        # all the products
                        product_list_li = product_list_ul.find_all("li")
                        if product_list_li:

                            # here we find the available products
                            div_PM_ASCriterionNbProduct = soup.select_one(
                                "div.PM_ASCriterionNbProduct"
                            )
                            if div_PM_ASCriterionNbProduct:
                                # get the number of th only available products
                                availableproducts = int(
                                    div_PM_ASCriterionNbProduct.getText().strip("()")
                                )
                                print(f"{availableproducts} products is available")

                                # get the href of the only available product
                                for i in range(availableproducts):
                                    if product_list_li[i].find(
                                        "div", class_="product-container"
                                    ):
                                        href = product_list_li[i].select_one(
                                            "div.product-container div link"
                                        )["href"]
                                        product_category_dict = {
                                            "parent_category": category_name,
                                            "name": product_list_li[i].select_one(
                                                "meta"
                                            )["content"],
                                            "link": href,
                                            "stock": -1,
                                        }
                                        category_list.append(product_category_dict)

        # if there is not any form tag
        # ====================================================================
        else:
            product_list_ul = soup.find("ul", id="product_list_cat")
            if product_list_ul:
                # all the products
                product_list_li = product_list_ul.find_all("li")
                if product_list_li:
                    for product in product_list_li:
                        if product.select_one(
                            "div.product-container div.button-container-sabad"
                        ):

                            product_category_dict = {
                                "parent_category": category_name,
                                "name": product.select_one("meta")["content"],
                                "link": href,
                                "stock": -1,
                            }
                            category_list.append(product_category_dict)


# Function to save data to Excel
def save_in_excel(category_list: list[dict]) -> str:
    """
    Save the category list to an Excel file.
    """
    wb = Workbook()
    sheet1 = wb.add_sheet("Sheet 1")
    excelfile: str = "xlwt_products.xls"
    row_index: int = 0
    column_index: int = 0

    for category in category_list:
        for value in category.values():
            sheet1.write(row_index, column_index, value)
            column_index += 1
        row_index += 1
        column_index = 0

    wb.save(excelfile)



def read_excel(excelfile: str, start_row: int = 0, end_row: int = 0) -> list[tuple]:
    """
    Read data from an Excel file and return it as a list of tuples (link, stock).

    Args:
        excelfile (str): Path to the Excel file.
        start_row (int): Row to start reading from (0-based).
        end_row (int | None): Row to stop reading at (exclusive). If None, reads until the end.

    Returns:
        list[tuple]: A list of (link, stock) pairs.
    """
    workbook = xlrd.open_workbook(excelfile)
    sheet = workbook.sheet_by_index(0)

    # If end_row is not specified, read until the last non-empty row
    if end_row is None:
        end_row = sheet.nrows  # Total rows in the sheet

    # Read columns C (index 2) and D (index 3) from start_row to end_row-1
    link = sheet.col_values(2, start_rowx=start_row, end_rowx=end_row)
    stock = sheet.col_values(3, start_rowx=start_row, end_rowx=end_row)

    return list(zip(link, stock))


async def go_to_stock_page(link_stock_list: list):
    return await run_requests(link_stock_list)


# Recursive function to traverse sub-categories
@timer
async def get_category_recursive(
    category_li: BeautifulSoup, category_name: str, category_list: list[dict]
):
    """
    Recursively fetch sub-categories and product links.
    Args:
        category_li (BeautifulSoup): The current category list item.
        category_name (str): The name of the current category.
        category_list (list[dict]): The list to store product information.
    """
    product_category_dict = {}

    if category_li:
        sub_category_ul = get_sub_category_ul(category_li)

        # if there is a category list
        if sub_category_ul:
            li_category_items = get_sub_categories_li(sub_category_ul)
            if li_category_items:

                # for each category in category_list
                for category_li in li_category_items:
                    subcategory_name = show_sub_categories(category_li)

                    # information about each category is saved a dictionary
                    product_category_dict = {
                        "parent_category": category_name,
                        "name": subcategory_name,
                        "link": "",
                        "stock": -1,
                    }
                    category_list.append(product_category_dict)

                    await get_category_recursive(
                        category_li=category_li,
                        category_name=subcategory_name,
                        category_list=category_list,
                    )

        # now we get the final product in the buttom subcategory
        elif category_li.find("a"):
            href = category_li.select_one("a")["href"]
            product_category_dict["link"] = href
            await get_product_links(
                href, category_name=category_name, category_list=category_list
            )


async def Read_Data_of_Site():
    """
    goes to the ickala for the first time and 
    starts to navigate the main menue in the ickala
    """
    url = "https://ickala.com/"
    product_category_dict = {}
    global category_list
    category_list = []

    soup = await get_page(url)
    category_ul = soup.select_one("div.block_content ul")
    if category_ul:
        li_category_items = get_sub_categories_li(category_ul)

        if li_category_items:
            for category_li in li_category_items:
                time.sleep(1)
                category_name = show_sub_categories(category_li)
                product_category_dict = {
                    "parent_category": "none_parent",
                    "name": category_name,
                    "link": "",
                    "stock": -1,
                }
                category_list.append(product_category_dict)
                await get_category_recursive(category_li, category_name, category_list)
            save_in_excel(category_list)


async def main_operation():
    """
    here we  ask user what he wants
    whether he wants to create a excel file
    if he wants to update all the excel file or just a part of it(just the ones which doesnt have quantity)
    """
    x: int = int(
        input(
            "if you want to enter data in an excel file enter 1 and if you want to update the stock enter 2: "
        )
    )
    if x == 1:
        print("you are going to enter data in an excel file")
        
        await Read_Data_of_Site()

    else:
        # here is the name of the excel file that is stored in the directory
        # of the python file
        excelfile = "xlwt_products.xls"
        if os.path.exists(excelfile):
            start_row: int = int(
                input(
                    """
                if you want to update the stock enter from the begining enter 1
                otherwise enter the numeber of the the desired row 

                """
                )
            )
            end_row: int = int(
                input(
                    """
                    if you want to update the stock enter until the end enter 1
                    otherwise enter the numeber of the the desired row 
                    """
                )
            )
            if start_row == 1:
                if end_row == 1:
                    print(
                        "you are going to update the stock from the begining until the end"
                    )
                    await update_stock(excelfile)
                else:
                    print(
                        f"you are going to update the stock from beginning until row {end_row}"
                    )
                    await update_stock(excelfile, start_row=start_row, end_row=end_row)

            else:
                if end_row == 1:
                    print(
                        f"you are going to update the stock from row {start_row} until the end"
                    )
                    await update_stock(excelfile, start_row=start_row)
                else:
                    # Update stock from row y to row z
                    if start_row < end_row:
                        print(
                            f"you are going to update the stock from row {start_row} until row {end_row}"
                        )
                        await update_stock(
                            excelfile, start_row=start_row, end_row=end_row
                        )
                    else:
                        # Invalid range, handle as needed
                        print(
                            f"Invalid range: start row {start_row} is greater than end row {end_row}. Please check your input."
                        )
                        await main_operation()
        else:
            print("this file does not exist please enter 1")
            await main_operation()


# Main function
async def main():
    """
    Main function to orchestrate the scraping process.
    It fetches the main page, extracts categories and sub-categories,
    and retrieves product information.
    """
    await main_operation()
    print("Operation successfully finished.")


# Run the script
if __name__ == "__main__":
    try:
        asyncio.run(main())
    except Exception as e:
        print(f"error is {e}")
