import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

from openpyxl import load_workbook  # conda install openpyxl


# https://samukweku.github.io/data-wrangling-blog/spreadsheet/python/pandas/openpyxl/2020/05/19/Access-Tables-In-Excel.html

def test_dataframe(df):
    """This method performs a set of tests on a dataframe to see if it has the right properties

    Note that it check not only the presence of columns but also their spelling

    Todo: add other tests if needed
    """

    # assert dataframe contain the required fields
    assert df.index.name == 'years', f"expected index.name 'years' not found"
    assert 'years' in df.columns, f"expected column 'years' not found"
    assert 'capex' in df.columns, f"expected column 'capex' not found"
    assert 'opex' in df.columns, f"expected column 'opex' not found"
    assert 'revenue' in df.columns, f"expected column 'revenue' not found"


def create_cashflow_dataframe(startyear=2000, lifecycle=11,
                              capex={'years': [2001, 2002],
                                     'values': [-5_000_000, -5_000_000]},
                              opex={'years': list(range(2003, 2011)),
                                    'values': 8 * [-300_000]},
                              revenue={'years': list(range(2003, 2011)),
                                       'values': 8 * [1_500_000]}):
    """This method returns a dataframe with 'years' as index and index_name and columns years, capex, opex and revenue.

    The method first spans up a list of years based on 'startyear' and 'lifecycle'
    Then it initialises the other required columns: 'capex', 'opex' and 'revenue'
    Next it places the years-values combinations of 'capex', 'opex' and 'revenue' on the right places in de dataframe

    The end result is a dataframe that shows the capital expenses for each year

    Todo: add residual value
    """

    df = pd.DataFrame()

    # create list of years using startyear and lifecycle and set years as index
    years = list(range(startyear, startyear + lifecycle))
    df['years'] = years
    df.index = years
    df.index.name = 'years'

    # initialise capex, opex and revenue as zero
    df['capex'] = 0
    df['opex'] = 0
    df['revenue'] = 0

    # add capex from input
    for year in capex['years']:
        df['capex'].loc[year] = capex['values'][capex['years'] == year]

    # add opex from input
    for year in opex['years']:
        df['opex'].loc[year] = opex['values'][opex['years'] == year]

    # add revenue from input
    for year in revenue['years']:
        df['revenue'].loc[year] = revenue['values'][revenue['years'] == year]

    # assert that dataframe adheres to prescribed standards
    test_dataframe(df)

    return df


def combine_cashflow_dataframes(dfs):
    """We assume that dfs is a list of dataframes that has a capex, opex and revenue column and years as index

    We add all years to a list 'years'. Next we determine the min and max year in that combined list.
    We create a new dataframe named df_combined that has all available years as a column and set as index.
    Next we step through the list of dataframes and add one by one all values to df_combined.

    Finally we return the combined dataframe.

    Todo: see if it is useful to also add a base year, so that you can calculate npvs to a give baseyear
    """

    # assert all dataframes contain the required fields
    for df in dfs:
        test_dataframe(df)

    years = []
    for df in dfs:
        years = years + df.years.tolist()

    min_year = min(years)
    max_year = max(years)

    new_years = list(range(min_year, max_year + 1))

    df_combined = pd.DataFrame()
    df_combined['years'] = new_years
    df_combined['capex'] = 0
    df_combined['opex'] = 0
    df_combined['revenue'] = 0
    df_combined.index = new_years
    df_combined.index.name = 'years'

    for df in dfs:
        for year in df.years.tolist():
            df_combined['capex'].loc[year] = df_combined['capex'].loc[year] + df['capex'].loc[year]
            df_combined['opex'].loc[year] = df_combined['opex'].loc[year] + df['opex'].loc[year]
            df_combined['revenue'].loc[year] = df_combined['revenue'].loc[year] + df['revenue'].loc[year]

    return df_combined


def calculate_npv(df, baseyear=2000, interest=0.07):
    """This method expects a dataframe that has years as index and index_name, and at least has columns
    named years, capex, opex, revenue.

    The method sums up all cashflows per year and adds these as a separate columns
    Also a cumulative cashflow column is added
    Next the npv is calculated
    Also a cumlative npv column is added

    Todo: see if it is useful to also add a 'baseyear', so that you can calculate npvs to a given 'baseyear'
    """

    # assert that dataframe adheres to prescribed standards
    test_dataframe(df)

    # collect the cashflows and add a 'cashflow' column
    df['cashflow'] = df.capex + df.opex + df.revenue

    # add the cumsum of cashflows to the 'cashflow_sum' column
    df['cashflow_sum'] = df['cashflow'].cumsum()

    # intitialise the 'npv' column with zeros
    df['npv'] = 0

    # calculate the npv through the years from the 2nd year up to the end and add the values to the 'npv' column
    for year in df.years.tolist()[1:]:
        # C_0 = C_n (1 + r) ** -n      (see Ports and Waterways - Part I - Equation 2.2, p. 42)
        df['npv'].loc[year] = df['cashflow'].loc[year] * (1 + interest) ** (-1 * (year - baseyear))

    # add the cumsum of npvs to the 'npv_sum' column
    df['npv_sum'] = df['npv'].cumsum()

    return df


def create_npv_plot(df, title=r'CAPEX, OPEX and Revenues and NPV', fname=r'test.png', x1=0, y1=0, x2=0, y2=0,
                    cash_flow_lims=[-1000, 1000], npv_lims=[-1000, 1000]):
    """This method creates a basic plot"""

    # assert that dataframe adheres to prescribed standards
    test_dataframe(df)

    # preset fontsize and legend fontsize
    fontsize = 20
    fontsize_legend = 15

    # initialise figure
    fig, ax1 = plt.subplots(1, 1, sharex=True, figsize=(16, 8))
    plt.rcParams['font.size'] = fontsize

    offset = 0.25
    width = 0.25

    plt.axis('off')

    ax1 = fig.add_subplot(1, 1, 1)

    # ----

    ax1.bar([x - 1 * offset for x in df['years']], height=list(df['capex'] / 10 ** 6), color='red', width=width,
            label='CAPEX')
    ax1.bar([x + 0 * offset for x in df['years']], height=list(df['opex'] / 10 ** 6), color='blue', width=width,
            label='OPEX')
    ax1.bar([x + 1 * offset for x in df['years']], height=list(df['revenue'] / 10 ** 6), color='green', width=width,
            label='Revenue')

    ax1.legend(loc='lower left', ncol=3, fontsize=fontsize_legend, bbox_to_anchor=(x1, y1), frameon=False)

    ax1.set_title(title, fontsize=fontsize)
    ax1.set_xlabel(r'Years', fontsize=fontsize, labelpad=20)
    ax1.set_ylabel(r'Cash flows ($10^6$ Euro)', fontsize=fontsize)
    ax1.set_xticks(np.arange(0, max(df['years']) + 1, 1))
    ax1.set_xticklabels(['{:.0f}'.format(x) for x in np.arange(0, max(df['years']) + 1, 1)], rotation=90,
                        fontsize=fontsize)

    ax1.grid(which='major', axis='both')
    ax1.set_xlim([df.years.min(), df.years.max() + 1])
    ax1.set_ylim(cash_flow_lims)

    # ----

    ax2 = ax1.twinx()  # instantiate a second axes that shares the same x-axis

    ax2.plot(list(df['years']), list(df['npv_sum'] / 10 ** 6), color='red', marker='o', label='NPV')

    ax2.legend(loc='lower right', fontsize=fontsize_legend, bbox_to_anchor=(x2 + 1, y2), frameon=False)

    ax2.set_ylabel('NPV ($10^6$ Euro)', fontsize=fontsize)  # we already handled the x-label with ax1

    ax2.set_ylim(npv_lims)  # NB: you want to take care that the y=0 of ax1 and ax2 align to avoid confusion

    # ----

    fig.tight_layout = True


def Inputs_2_cashflow(Inputs,
                      startyear=2000,
                      lifecycle=11,
                      subsystem='Wind energy source & Transport',
                      element='Offshore wind park',
                      component='Foundations',
                      cashflow_categories=['Development and Project Management', 'Procurement', 'Installation'],
                      Debug=False):
    """
    Assuming columns Sub-system, Element and Component allways exist

    Method returns cashflow dataframe
    """

    Number_of_units = Inputs[
        (Inputs['Sub-system'] == subsystem) &
        (Inputs['Element'] == element) &
        (Inputs['Component'] == component) &
        (Inputs['Category'] == 'General') &
        (Inputs['Description'] == 'Number of units')].Number.item()

    if Debug:
        display('Construction items {}: {} units'.format(component, Number_of_units))

    Capex_component = Inputs[
        (Inputs['Sub-system'] == subsystem) &
        (Inputs['Element'] == element) &
        (Inputs['Component'] == component) &
        (Inputs['Category'].isin(cashflow_categories))].Number.sum() * Number_of_units

    if Debug:
        display('CAPEX component {}: {} eu for {} unit(s)'.format(component, Capex_component, Number_of_units))

    Construction_duration = Inputs[
        (Inputs['Sub-system'] == subsystem) &
        (Inputs['Element'] == element) &
        (Inputs['Component'] == component) &
        (Inputs['Category'] == 'Capex') &
        (Inputs['Description'] == 'Construction duration')].Number.item()
    if Debug:
        display('Construction duration {}: {} years'.format(component, Construction_duration))

    Construction_allocation = Inputs[
        (Inputs['Sub-system'] == subsystem) &
        (Inputs['Element'] == element) &
        (Inputs['Component'] == component) &
        (Inputs['Category'] == 'Capex') &
        (Inputs['Description'] == 'Capex allocation')].Number.item().split(",")
    Construction_allocation = [float(x) for x in Construction_allocation]

    if Debug:
        display('Construction allocation {}: {} per year'.format(component, Construction_allocation))

    Opex_component = Inputs[
                         (Inputs['Sub-system'] == subsystem) &
                         (Inputs['Element'] == element) &
                         (Inputs['Component'] == component) &
                         (Inputs['Category'] == 'Opex')].Number.item() / 100 * Capex_component

    if Debug:
        display('OPEX component {}: {} eu for {} unit(s)'.format(component, Opex_component, Number_of_units))

    Revenue_component = 0

    if Debug:
        display('Revenue {}: {} euro/unit'.format(component, Revenue_component))

    df = create_cashflow_dataframe(startyear=startyear, lifecycle=lifecycle,
                                   capex={'years': list(range(startyear + 1, startyear + 1 + (Construction_duration))),
                                          'values': [-item * Capex_component for item in Construction_allocation]},
                                   opex={'years': list(
                                       range(startyear + 1 + (Construction_duration), startyear + lifecycle)),
                                       'values': len(list(
                                           range(startyear + 1 + (Construction_duration), startyear + lifecycle))) * [
                                                     -Opex_component]},
                                   revenue={'years': list(
                                       range(startyear + 1 + (Construction_duration), startyear + lifecycle)),
                                       'values': len(list(range(startyear + 1 + (Construction_duration),
                                                                startyear + lifecycle))) * [Revenue_component]})
    return df
