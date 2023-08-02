import pandas as pd

#Opening correct file and required sheets
non_asthma = pd.read_excel('Asthma Results - Raw.xlsx',sheet_name='Non-Asthmatic')
asthmatic = pd.read_excel('Asthma Results - Raw.xlsx',sheet_name='Asthmatic')

#function definition
def correct_outliers():
    #setting the quartiles and medians for each feature
    non_asthmatic_quartiles = non_asthma.quantile([0.25,0.75])
    non_asthma_medians = non_asthma.median()

    asthmatic_quartiles = asthmatic.quantile([0.25,0.75])
    asthmatic_medians = asthmatic.median()

    #iterating through each necessary feature and updating the values that lie outside quartile 1 and 3.
    for feature in non_asthma.columns[2:82]:
        Q_1 = non_asthmatic_quartiles.loc[0.25,feature]
        Q_3 = non_asthmatic_quartiles.loc[0.75,feature]
        median = non_asthma_medians[feature]

        for i, x in enumerate(non_asthma[feature]):
            if x < Q_1 or x > Q_3:
                non_asthma.at[i,feature] = median

    for feature in asthmatic.columns[2:]:
        Q_1 = asthmatic_quartiles.loc[0.25,feature]
        Q_3 = asthmatic_quartiles.loc[0.75,feature]
        median = asthmatic_medians[feature]

        for i, x in enumerate(asthmatic[feature]):
            if x < Q_1 or x > Q_3:
                asthmatic.at[i,feature] = median 

    #saving results in a new Excel file.
    non_asthma.to_excel('Corrected Data.xlsx', index=False,sheet_name='Non-Asthmatic_Corrected')

    with pd.ExcelWriter('Corrected Data.xlsx',engine='openpyxl',mode = 'a') as writer:
        asthmatic.to_excel(writer, index=False,sheet_name='Asthmatic_Corrected')




correct_outliers()