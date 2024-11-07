# Ben Tunney
# DATA Initiative Part 2 Task

# resources
# https://www.scrapingbee.com/blog/selenium-python/#locating-elements

# web scraping
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup

# data management
import pandas as pd
import csv
import openpyxl

def driver(path):
    """
    set up selenium webdriver, output html for given search
    """

    # headless Chrome setup
    options = Options()
    options.headless = True
    options.add_argument("--window-size=1920,1200")

    # setup Chrome with the specified options
    driver = webdriver.Chrome(options=options, executable_path="chromedriver-mac-x64/chromedriver")

    # navigate to the  website
    driver.get(path)

    # Output the page source to the console
    html = driver.page_source

    # Close the browser session cleanly to free up system resources
    driver.quit()

    return html


def scrape_pages(csv_filename, start_page, end_page):
    """
    scrape range of specified amazon search pages, update csv file
    """

    # open csv file into list of lists
    product_review_list = []
    with open(csv_filename, newline='', encoding='utf-8') as csvfile:
        csvreader = csv.reader(csvfile)
        for row in csvreader:
            product_review_list.append(row)

    # scrape reviews for each specified product search page
    for page_num in range(start_page, end_page+1):

        search_page_fname = f"https://www.amazon.com/s?i=computers&amp;rh=n%3A565108%2Cp_36%3A2421886011&amp;s=exact-aware-popularity-rank&amp;page={page_num}&amp;content-id=amzn1.sym.4d915fa8-ca05-4073-b385-a93e1e1dd22d&amp;pd_rd_r=a42c9ae5-5a10-4802-8c2f-5f25ef0e364a&amp;pd_rd_w=aoUV0&amp;pd_rd_wg=9r6MB&amp;pf_rd_p=4d915fa8-ca05-4073-b385-a93e1e1dd22d&amp;pf_rd_r=B31SVVTK8X16Q535JREG&amp;qid=1719941120&amp;ref=sr_pg_{page_num}"

        product_review_list = scrape_search_page(search_page_fname, product_review_list,page_num)

    # create DataFrame
    df = pd.DataFrame(product_review_list[1:], columns=product_review_list[0])

    # write the DataFrame to a CSV file
    df.to_csv(csv_filename, index=False)


def scrape_search_page(search_page_fname, product_review_list, page_num):
    """
    scrape reviews for all products on search page, add them to product_review_list
    """

    # product search page
    html = driver(search_page_fname)

    # get all product search results
    soup = BeautifulSoup(html, 'html.parser')
    search_results = soup.find_all('div',
                    class_='sg-col-4-of-24 sg-col-4-of-12 s-result-item s-asin sg-col-4-of-16 sg-col s-widget-spacing-small sg-col-4-of-20 gsx-ies-anchor')


    # for each product on page
    for product_idx in range(len(search_results)):

        print(f"Scraping reviews for search page #{page_num}, product #{product_idx}.")

        # get product url
        href = search_results[product_idx].find('a', href=True)['href']
        product_path = "https://www.amazon.com" + href.split("ref")[0]

        # iterate through first 10 reviews (NOTE: due to scraping amazon limitations,
        # only first 10 pages of reviews can be accessed)
        page_idx = 1
        while page_idx < 11:

            # attempt to access page of reviews
            try:
                review_page = product_path + f"ref=cm_cr_getr_d_paging_btm_next_{page_idx}?ie=UTF8&reviewerType=all_reviews&pageNumber={page_idx}"
                review_page = review_page.replace("dp", "product-reviews")
                review_page_html = driver(review_page)
                review_page_soup = BeautifulSoup(review_page_html, 'html.parser')
                results = review_page_soup.find_all('div', {'data-hook': 'review'})

                # get each review
                for review_idx in range(len(results)):
                    review = results[review_idx]

                    # attempt to get metadata per review
                    try:
                        product_identifier = review_page.split("/")[-2]
                    except:
                        product_identifier = None
                    try:
                        product_name = review_page.split("/")[-4]
                    except:
                        product_name = None
                    try:
                        title = review.find('a', {'data-hook': 'review-title'}).text.strip().split("\n")[1]
                    except:
                        title = None
                    try:
                        rating = review.find('i', {'data-hook': 'review-star-rating'}).text.strip()
                    except:
                        rating = None
                    try:
                        date = review.find('span', {'data-hook': 'review-date'}).text.strip()
                    except:
                        date = None
                    try:
                        review_text = review.find('span', {'data-hook': 'review-body'}).text.strip()
                    except:
                        review_text = None
                    try:
                        reviewer = review.find('span', {'class': 'a-profile-name'}).text.strip()
                    except:
                        reviewer = None
                    try:
                        helpful_count = review.find('span', {'data-hook': 'helpful-vote-statement'}).text.strip()
                    except:
                        helpful_count = None
                    try:
                        vp = review.find('span', {'data-hook': 'avp-badge'}).text.strip()
                    except:
                        vp = None

                    # store metadata into lists
                    product_review_list.append([product_identifier, product_name, title, date, review_text, rating,
                                                reviewer, helpful_count, vp, page_idx, page_num])

            # break if cannot access review page url
            except:
                break

            page_idx += 1

    return product_review_list


def main():

    # code to create initial csv file
    """
    # Define the filename
    filename = 'amazon_reviews.csv'

    product_review_list = [["product_id", "product_name", "review_title", "review_date", "review_text",
                            "review_star_rating", "reviewer", "helpful_votes", "verified", "review_page_num",
                            "search_page_num"]]

    # Create the DataFrame
    df = pd.DataFrame(columns=product_review_list)

    # Write the DataFrame to a CSV file
    df.to_csv(filename, index=False)
    """

    # scrape pages of interest
    # scrape_pages("amazon_reviews.csv", 8, 8)

    # clean newlines from csv
    """
    df = pd.read_csv("amazon_reviews.csv")
    df['review_text'] = [str(x).replace('\n', '') for x in df['review_text']]

    # write the cleaned DataFrame to a CSV file
    df.to_csv("amazon_reviews1.csv", index=False)
    """
    read_file = pd.read_csv('amazon_reviews1.csv')
    read_file.to_excel('Reviews-Ben_Tunney.xlsx', index=None, header=True)


if __name__ == '__main__':
    main()

