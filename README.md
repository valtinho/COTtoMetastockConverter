# COT to Metastock Converter
This tool extracts the S&amp;P 500 data from the Commitment of Traders (COT) Futures-Only report so it can be charted with the Metastock application. Net positions are calculated for the 3 main groups: large traders (non-commercials), commercials, and small speculators (non-reportables). In addition, the spread between the net positions of the commercials versus the small speculators is calculated. Column header names and ticker symbols can be modified in the "configuration.txt" file if the CFTC decides to change the format of the Excel file in the future.

To download the latest COT report from the U.S. Commodity Futures Trading Commission's (CFTC) website, please go to:

<a href="http://www.cftc.gov/MarketReports/CommitmentsofTraders/HistoricalCompressed/index.htm" target="_blank">http://www.cftc.gov/MarketReports/CommitmentsofTraders/HistoricalCompressed/index.htm</a>

Download the Excel version as a zip file. Extract its contents and supply the xls file to the COT to Metastock Converter Application.

Screenshot of the COT to Metastock Converter Application:
<img src="https://raw.githubusercontent.com/valtinho/COTtoMetastockConverter/master/COTtoMetastockConverter_Screenshot1.png" alt="COT to Metastock Converted Application" />

I have chosen the legacy report (not disaggregated), because it is the one with the longest history. I also chose the futures-only report (without options data), because it is easier to interpret and make trading assumptions from. Please note, however, that neither COT, nor any other indicator, should not be used alone for trading decisions.

COT, like anything else, is full of noise most of the time. But, in the 1970's, the infamous Larry Williams discovered that it was worth considering it when the disparity between the smart-money (commercials) and the dumb-money (small speculators) were at extremes. If large traders tend to be trend-followers, they are usually late to the game and are not worth considering for leading assumptions. 

Following Larry William's basic theory, the spread is calculated from the net positions (of longs and shorts) of the commercials divided the net positions of the small speculators. The "extremes" can be spotted with standard Bollinger Bands applied to the data. Bollinger bands make normal distribution assumptions based on sample sizes that are usually too small to have any statistical validity (less than 30), but they do serve as a very rough guide. Ultimately, we can probably say that if the COT spread line is above or below the Bollinger Bands, then an extreme value is being observed.

The ideal bullish scenario is to observe S&P 500 prices declining rapidly, while the COT spread line is shooting up above the upper Bollinger Band. This means that the commercials are extremely net long, while the small speculators are extremely net short. Bearish scenarios would be the exact opposite.

It is worth noting that in the last few years, the dumb money has been right as often as the smart money. So, perhaps the small speculators as not as dumb as most would assume, or simply that the S&P trend is so strong that commercials have lost their power of influence. Only time will tell...

Here is a screenshot of the data plotted in Metastock. The red-dotted lines are the Bollinger Bands.

<img src="https://raw.githubusercontent.com/valtinho/COTtoMetastockConverter/master/Metastock_Screenshot.png" alt="Metastock Screenshot" />
