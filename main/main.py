from webscraping.extract_data_flex_monitor import WebScraping

if __name__ == "__main__":
    user = "username"  # Replace with actual username
    password = "password"  # Replace with actual password
    scraper = WebScraping()
    scraper.web_scraping(user, password)
