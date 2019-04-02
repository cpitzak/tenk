from sqlalchemy import Column, DateTime, String, Integer, Float, ForeignKey, func
from sqlalchemy.orm import relationship, backref
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker

Base = declarative_base()


class Metrics(Base):

    __tablename__ = 'metrics'
    id = Column(Integer, primary_key=True)
    ticker = Column(String)
    report_date = Column(String)
    gross_profit_margin = Column(Float)
    sga_gross_profit = Column(Float)
    rd_gross_profit = Column(Float)
    depreciation_gross_profit = Column(Float)
    interest_expense_operating_income = Column(Float)
    taxes = Column(Float)
    net_earnings = Column(Float)
    net_earnings_by_revenue = Column(Float)
    eps_basic = Column(Float)
    eps_diluted = Column(Float)
    return_on_assets = Column(Float)
    liabilities_by_equity = Column(Float)
    retained_earnings = Column(Float)
    return_on_shareholder_equity = Column(Float)
    capital_expend = Column(Float)
    net_stock_buy_back = Column(Float)


# class KeyRatios(Base):
#
#     __table__ = 'key_ratios'
#     id = Column(Integer, primary_key=True)
#     ticker = Column(String)
#     report_date = Column(String)
#     revenue = Column(Float)
#     gross_margin = Column(Float)
#     net_income = Column(Float)
#     operating_margin = Column(Float)
#     eps = Column(Float)
#     dividends = Column(Float)
#     payout_ratio = Column(Float)
#     share_mil = Column(Float)
#     book_value_per_share = Column(Float)


engine = create_engine('sqlite:///buffett_calcs.sqlite')
# engine = create_engine('postgresql://postgres:password@localhost:5432/tenkdb')

# export sql file:
# postgresql-8.1.11.tar\postgresql-8.1.11\src\bin\pg_dump -h localhost -p 5432 -U postgres tenkdb > c:\data\metrics.sql

session = sessionmaker()
session.configure(bind=engine)
Base.metadata.create_all(engine)