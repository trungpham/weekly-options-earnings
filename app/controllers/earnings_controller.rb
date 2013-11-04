require 'csv'
require 'open-uri'
require 'nokogiri'

class EarningsController < ApplicationController

  EARNINGS_CALENDAR_URL = 'http://biz.yahoo.com/research/earncal/#{date}.html' #"http://biz.yahoo.com/research/earncal/20131104.html"
  WEEKLY_OPTIONS_URL = "http://www.cboe.com/publish/weelkysmf/weeklysmf.xls"
  UPCOMING_EARNINGS_URL = "http://finviz.com/export.ashx?v=111&f=cap_largeover,earningsdate_nextdays5,sh_avgvol_o1000,sh_opt_option,sh_price_o30&ft=4"

  def index

    stocks_with_weekly = Rails.cache.fetch('weekly_options', expires_in: 5.minutes) do
      weekly_options_ss = Roo::Excel.new(WEEKLY_OPTIONS_URL)
      lastest_list_date = get_latest_list_date(weekly_options_ss)
      result = []
      (weekly_options_ss.first_row.to_i..weekly_options_ss.last_row.to_i).each do |row|
        if (weekly_options_ss.cell(row, 'D') == 'Equity' && weekly_options_ss.cell(row, 'E') == lastest_list_date)
          result << weekly_options_ss.cell(row, 'A')
        end
      end

      result
    end

    stocks_with_upcoming_earnings = Rails.cache.fetch('upcoming_earnings', expires_in: 5.minutes) do

      open(UPCOMING_EARNINGS_URL) do |file|

        csv = CSV.new(file.string, :headers => true, :header_converters => :symbol)
        result = csv.to_a.map {|row| row.to_hash }

      end


    end

    @stocks = add_earning_date(get_stocks_with_weekly_and_earnings(stocks_with_weekly, stocks_with_upcoming_earnings))
  end

  private
  def get_latest_list_date(weekly_options_ss)
    list_date = nil
    (1..10).each do |row|
      if weekly_options_ss.sheet(0).cell(row, 'E') == 'List Date'
        list_date = weekly_options_ss.sheet(0).cell(row+1, 'E')
        break
      end
    end
    list_date
  end

  #earnings:
  #[{:no=>"46", :ticker=>"WFM", :company=>"Whole Foods Market, Inc.", :sector=>"Services", :industry=>"Grocery Stores", :country=>"USA", :market_cap=>"23509.62", :pe=>"43.66", :price=>"63.30", :change=>"0.27%", :volume=>"2798675"}]
  #weekly: ['A', 'B', 'C']
  def get_stocks_with_weekly_and_earnings(weekly, earnings)
    result = {}
    earnings.each do |earning|
      if weekly.include?(earning[:ticker])
        result[earning[:ticker]] = earning
      end
    end
    result
  end

  def add_earning_date(stocks)
    earning_dates = Rails.cache.fetch('earning_dates', expires_in: 24.hours) do
      current_date = Date.today
      stocks_and_date = {}
      (0..5).each do |i|
        date_value = current_date.strftime('%m-%d-%Y')
        doc = Nokogiri::HTML(open(EARNINGS_CALENDAR_URL.sub('#{date}', current_date.strftime('%Y%m%d'))))
        doc.css('tr td a[href^="http://finance.yahoo.com/q?s="]').each do |node|
          stocks_and_date[node.text] = date_value
        end
        current_date = current_date.next_day
      end
      stocks_and_date
    end

    stocks.each do |k, v|
      v[:er_date] = earning_dates[k]
    end
    stocks
  end
end