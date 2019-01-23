import numpy as np
import matplotlib.pyplot as plt
import pandas as pd

##handle version issue on my mac
plt.interactive(True)

##read data into DataFrame
def generateDatabaseFromExcel(file_name,sheet_names):

    #Read from excel
    dataFrames = {}
    for sheet_name in sheet_names:
        dataFrames[sheet_name] = pd.read_excel(file_name, sheet_name)

    # Data Normalization
    for dfName in dataFrames:
        df = dataFrames[dfName]
        
        ID = df["ID"]
        print(ID)
        dataFrames[dfName] = (df - df.min()) / (df.max() - df.min())
        dataFrames[dfName]["ID"] = ID

    # Add Year info

    for dfName in dataFrames:
        dataFrames[dfName]['Year'] = int(dfName[-2:])

        # Combine the dataframes into a database
        database = pd.concat([dataFrames[i] for i in dataFrames], axis=0, sort=True)
        
    #print(database)

    return database



def distributionHistogram(database,subject,Year):
    data = database[database.Year == Year][subject]
    fig = plt.figure()
    ax = fig.add_subplot(1, 1, 1)
    ax.boxplot(data)
    fig.show()
    #fig.savefig("xxx.png")

def distributionBoxplot(database,subject,Year):
    data = database[database.Year == Year][subject]
    fig = plt.figure()
    ax = fig.add_subplot(1, 1, 1)
    ax.hist(data, bins=10, color="blue", alpha=0.7)
    fig.show()
    #fig.savefig("xxx.png")

def medianTrend(database,subject):
    database.groupby("Year")[subject].median().plot(legend = True)


def meanTrend(database,subject):
    database.groupby("Year")[subject].mean().plot()


def trendOverYears(database,subject):
    medians = database.groupby("Year")[subject].median().tolist()
    means = database.groupby("Year")[subject].mean().tolist()
    years = database.groupby("Year")["Year"].median().tolist()
    plt.plot(years,means,"o-")
    plt.plot(years,medians,"o-")
    plt.title(subject+" Trend Over Years")
    plt.legend(["Mean Trend","Median Trend"],loc = "upper right")
    plt.xticks(years)
    plt.show()
    # plt.savefig("xxx.png")






def main():
    sheet_names = ["Year10", "Year11", "Year12"]
    file_name = "SUTD.xlsx"
    database = generateDatabaseFromExcel(file_name,sheet_names)

    subject = "B"
    Year = 10
    # distributionHistogram(database,subject,Year)
    # distributionBoxplot(database, subject, Year)
    # medianTrend(database,subject)
    # medianTrend(database,"L")
    # meanTrend(database, subject)
    trendOverYears(database,subject)



    # print(database.groupby("Year")["Year"].median().tolist())



if __name__ == '__main__':
    main()
