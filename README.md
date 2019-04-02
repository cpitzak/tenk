# Tenk

This program reads 10k financial statements (balance sheet, income statement, and cash flow statement) and computes the metrics found in [Warren Buffett and the Interpretation of Financial Statements: The Search for the Company with a Durable Competitive Advantage](https://www.amazon.com/Warren-Buffett-Interpretation-Financial-Statements/dp/1416573186)

This program use to read a file with your tickers and then connect to morningstar and download the financial statements for you. However, morningstar no longer allows that see: [https://gist.github.com/hahnicity/45323026693cdde6a116](https://gist.github.com/hahnicity/45323026693cdde6a116) . I left the morningstar downloader code I wrote in the morningstar_downloader.py file encase they update this.

# Setup
1. conda create -n tenk python=3.6 anaconda
2. Activate your conda environment:
    1. Windows: activate tenk
    2. Linux: source activate tenk
3. pip install sqlalchemy==1.3.1

# Download 10k statements
1. Go to the links below and click the csv export icon:
    1. Income Statement: http://financials.morningstar.com/income-statement/is.html?t=AAPL&region=usa&culture=en-US
    2. Balance Sheet: http://financials.morningstar.com/balance-sheet/bs.html?t=AAPL&region=usa&culture=en-US
    3. Cash flow: http://financials.morningstar.com/cash-flow/cf.html?t=AAPL&region=usa&culture=en-US
2. Place the downloaded files in the **financials** directory
3. Repeat this process with as many tickers as you want. Note:
    1. Do not rename the filenames that came from morningstar
    2. Make sure to download Income, Balance, and Cash flow statement for each ticker

# Run the program
`python runner.py`

# View the output
1. Go into **buffett_calcs** directory and open the excel files you are interested in viewing
2. Connect to the sqlite database **buffett_calcs.sqlite** if you want to view the output in sql

# Demo
1. As a demo I included AAPL and MSFT balance, income, and cash flow statements in the **financials** directory, the excel output in the **buffett_calcs** directory, and the sql db (if you want to view in sql instead of excel) at the root.