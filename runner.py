import csv
import os

from morningstar.morningstar_downloader import MorningstarDownloader
from morningstar.morningstar_parser import MorningstarParser

if __name__ == '__main__':
    # Code commented out because morningstar no longer allows this, for more information and to see:
    # https://gist.github.com/hahnicity/45323026693cdde6a116
    # tickers = []
    # first_row = True
    # input_dir = 'input'
    # with open(os.path.join(input_dir, 'example_tickers.csv')) as f:
    #     r = csv.reader(f, delimiter=',')
    #     for row in r:
    #         if first_row:
    #             first_row = False
    #         else:
    #             tickers.append(row[0])
    # md = MorningstarDownloader(output_dir='financials')
    # md.download_tickers(tickers=tickers)
    # print('finished downloads')

    # morningstar naming for files
    BALANCE_SHEET = "Balance Sheet.csv"
    CASH_FLOW = "Cash Flow.csv"
    INCOME_STATEMENT = "Income Statement.csv"

    input_dir = 'financials'
    output_dir = 'buffett_calcs'
    years_of_data = 5

    files = [f for f in os.listdir(input_dir) if os.path.isfile(os.path.join(input_dir, f))]
    tickers = set([t.split(' ')[0] for t in files])

    mp = MorningstarParser()
    for ticker in tickers:
        ticker_files = [i for i in files if i.startswith(ticker + ' ')]
        income_statement = ''
        balance_sheet = ''
        cash_flow = ''
        for tfile in ticker_files:
            if tfile.endswith(BALANCE_SHEET):
                balance_sheet = os.path.join(input_dir, tfile)
            elif tfile.endswith(INCOME_STATEMENT):
                income_statement = os.path.join(input_dir, tfile)
            elif tfile.endswith(CASH_FLOW):
                cash_flow = os.path.join(input_dir, tfile)
        mp.process_morningstar_data(income_statement=income_statement,
                                    balance_sheet=balance_sheet, cash_flow=cash_flow,
                                    output_dir=output_dir, years=years_of_data, ticker=ticker)
        print('wrote {ticker} to {output_dir} directory'.format(ticker=ticker, output_dir=output_dir))
    print('done')
