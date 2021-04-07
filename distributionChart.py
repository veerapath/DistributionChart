import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from matplotlib.ticker import PercentFormatter

def generateDailyReturn(df):
    #Daily return(T+1) = [ ( ClosePrice(T+1) – ClosePrice(T) ) / ClosePrice(T) ] * 100
    result = (df['Bid Close'] - df['Bid Close'].shift(-1))/df['Bid Close'].shift(-1) * 100
    return result 

def plotHistogram(df, name):
    n_bins = (int)((df.max() - df.min())/0.05)

    plt.hist(df, bins=n_bins)
    plt.title(name)

    #enable y axis grid
    ax = plt.gca()
    ax.yaxis.grid(True)
    #set % format and decimal number
    ax.xaxis.set_major_formatter(PercentFormatter(100, decimals=2))

    #plt.xticks(np.arange(min(df['%Daily Return'].to_numpy()), max(df['%Daily Return'].to_numpy())+0.5, 1))
    plt.show()

def main():
    # read dataframes from excel file
    # use openpyxl instead of default xlrd engine which is no longer support for xlsx file
    dfs = pd.read_excel('PreciousMetalSpot.xlsx', sheet_name=None, engine='openpyxl')

    for sheet in dfs:
        #Access all precious metal data
        df = dfs[sheet]
        #generate daily return percent data
        dailyReturn = generateDailyReturn(df)
        #generate chart
        plotHistogram(dailyReturn,sheet.replace(' Spot',''))

if __name__ == '__main__':
    main()
