from webscraping.extract_data_flex_monitor import WebScraping
from dataprocessing.process_data import DataProcessing
from dataprocessing.prepare_file import FilePreparation

from pathlib import Path


if __name__ == "__main__":
    # <--- Instance classes --->
    scraper = WebScraping()
    data_processor = DataProcessing()
    prepare = FilePreparation()
    
    # <--- Variables --->
    pre_data = Path("data", "raw", "MonitorFlexExportacao.xls")
    source_data = Path("data", "raw", "MonitorFlexExportacao.xlsx")
    user = "username"  # Replace with actual username
    password = "password"  # Replace with actual password
        
    # <--- Calling functions --->
    scraper.web_scraping(user, password)
    prepare.convert_file(pre_data)
    data_processor.execute_code(source_data)
