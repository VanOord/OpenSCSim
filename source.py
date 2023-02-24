import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

from copy import deepcopy
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


def create_cashflow_dataframe(escalation_base_year=2023, lifecycle=50,
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
    years = list(range(escalation_base_year, escalation_base_year + lifecycle + 1))
    df['years'] = years
    df.index = years
    df.index.name = 'years'

    # initialise capex, opex and revenue as zero
    df['capex'] = 0
    df['opex'] = 0
    df['revenue'] = 0

    # add capex from input
    for index, year in enumerate(capex['years']):
        df.loc[year, 'capex'] = capex['values'][index]

    # add opex from input
    for index, year in enumerate(opex['years']):
        df.loc[year, 'opex'] = opex['values'][index]

    # add revenue from input
    for index, year in enumerate(revenue['years']):
        df.loc[year, 'revenue'] = revenue['values'][index]

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
            df_combined.loc[year, 'capex'] = (df_combined['capex'].loc[year] + df['capex'].loc[year]).copy()
            df_combined.loc[year, 'opex'] = (df_combined['opex'].loc[year] + df['opex'].loc[year]).copy()
            df_combined.loc[year,'revenue'] = (df_combined['revenue'].loc[year] + df['revenue'].loc[year]).copy()

    return df_combined


def calculate_npv(df, baseyear=2000, WACC=0.07):
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
    df['cashflow'] = df.capex.copy() + df.opex.copy() + df.revenue.copy()

    # add the cumsum of cashflows to the 'cashflow_sum' column
    df['cashflow_sum'] = df['cashflow'].cumsum()

    # intitialise the 'npv' column with zeros
    df['npv'] = 0

    # calculate the npv through the years from the 2nd year up to the end and add the values to the 'npv' column
    for year in df.years.tolist()[:]:
        df.loc[year, 'npv'] = df['cashflow'].loc[year] * (1 /((1 + WACC) ** (year - baseyear + 1)))

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
    ax1.set_xlim([df.years.min() - 1, df.years.max() + 1])
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
                      capex_categories=['Development and Project Management', 'Procurement',
                                        'Installation and Commissioning'],
                      opex_categories=['Yearly Variable Costs Rate', 'Insurance Rate'],
                      Debug=False):
    """
    Assuming columns Sub-system, Element and Component allways exist

    Method returns cashflow dataframe
    """

    # Escalation base year
    try:
        escalation_base_year = Inputs[
            (Inputs['Category'] == 'System input') &
            (Inputs['Description'].str.contains('Escalation base year'))].Number.item()

    except:
        escalation_base_year = 2000

    if Debug:
        display('Escalation base year {}: {}'.format(component, escalation_base_year))

    assert escalation_base_year <= startyear, f"escalation_base_year should be smaller or equal to startyear"

    # Escalation rate
    try:
        escalation_rate = Inputs[
            (Inputs['Category'] == 'System input') &
            (Inputs['Description'].str.contains('Escalation rate'))].Number.item()

    except:
        escalation_rate = 0.02

    if Debug:
        display('Escalation rate {}: {}'.format(component, escalation_rate))

    # Number of units
    try:
        Number_of_units = Inputs[
            (Inputs['Sub-system'] == subsystem) &
            (Inputs['Element'] == element) &
            (Inputs['Component'] == component) &
            (Inputs['Description'] == 'Number of units')].Number.item()
        Units = Inputs[
            (Inputs['Sub-system'] == subsystem) &
            (Inputs['Element'] == element) &
            (Inputs['Component'] == component) &
            (Inputs['Description'] == 'Number of units')].Unit.item()

    except:
        Number_of_units = 1

    if Debug:
        display('Number of units {}: {} {}'.format(component, Number_of_units, Units))

    # Construction duration (must be an integer)
    try:
        Construction_duration = int(Inputs[
                                        (Inputs['Sub-system'] == subsystem) &
                                        (Inputs['Element'] == element) &
                                        (Inputs['Component'] == component) &
                                        (Inputs['Description'] == 'Construction duration')].Number.item())
    except:
        Construction_duration = 3

    if Debug:
        display('Construction duration {}: {} years'.format(component, Construction_duration))

    assert isinstance(Construction_duration, int), f"Construction_duration must be an integer"

    # Share of Investments
    try:
        # isolate the rows that contain 'Share of Investments in Year' and remove the string to only get the year numbers
        years = list(Inputs[
            (Inputs['Sub-system'] == subsystem) &
            (Inputs['Element'] == element) &
            (Inputs['Component'] == component) &
            (Inputs['Description'].str.contains('Share of Investments in Year'))].Description.str.replace(
            'Share of Investments in Year ', ''))

        # transform the year numbers from string to in and sort them to be certain of the order
        years = [int(x) for x in years]
        years.sort()

        # now extract the allocations, since years are sorted we know for sure now that the allocations are in the right order
        Construction_allocation = []
        for year in years:
            Construction_allocation.append(Inputs[
                                               (Inputs['Sub-system'] == subsystem) &
                                               (Inputs['Element'] == element) &
                                               (Inputs['Component'] == component) &
                                               (Inputs['Description'].str.contains(
                                                   'Share of Investments in Year ' + str(year)))
                                               ].Number.item())
    except:
        Construction_allocation = [0.4, 0.3, 0.3]

    if Debug:
        display('Construction allocation {}: {} per year'.format(component, Construction_allocation))

    assert len(Construction_allocation)==Construction_duration, f"Length of Construction_allocation list must be equal to Construction_duration"

    # Economic Lifetime (must be an integer)
    try:
        Economic_lifetime = int(Inputs[
                            (Inputs['Sub-system'] == subsystem) &
                            (Inputs['Element'] == element) &
                            (Inputs['Component'] == component) &
                            (Inputs['Description'] == 'Economic Lifetime')].Number.item())

    except:
        lifecycle = 50

    if Debug:
        display('Economic Lifetime {}: {} years'.format(component, Economic_lifetime))

    # Depreciation Flag
    try:
        depreciation_flag = Inputs[
            (Inputs['Sub-system'] == subsystem) &
            (Inputs['Element'] == element) &
            (Inputs['Component'] == component) &
            (Inputs['Description'] == 'Depreciation Flag')].Number.item()

    except:
        depreciation_flag = 1

    if Debug:
        display('Depreciation Flag {}: {}'.format(component, depreciation_flag))

    # Yearly Variable Costs Flag
    try:
        yearly_variable_costs_flag = Inputs[
            (Inputs['Sub-system'] == subsystem) &
            (Inputs['Element'] == element) &
            (Inputs['Component'] == component) &
            (Inputs['Description'] == 'Yearly Variable Costs Flag')].Number.item()

    except:
        yearly_variable_costs_flag = 1

    if Debug:
        display('Yearly Variable Costs Flag {}: {}'.format(component, yearly_variable_costs_flag))

    # Yearly Variable Costs Rate
    try:
        yearly_variable_cost_rate = Inputs[
            (Inputs['Sub-system'] == subsystem) &
            (Inputs['Element'] == element) &
            (Inputs['Component'] == component) &
            (Inputs['Description'] == 'Yearly Variable Costs Rate')].Number.item()

    except:
        yearly_variable_cost_rate = 1

    if Debug:
        display('Yearly Variable Costs Rate {}: {}'.format(component, yearly_variable_cost_rate))

    # Insurance Flag
    try:
        insurance_flag = Inputs[
            (Inputs['Sub-system'] == subsystem) &
            (Inputs['Element'] == element) &
            (Inputs['Component'] == component) &
            (Inputs['Description'] == 'Insurance Flag')].Number.item()

    except:
        insurance_flag = 1

    if Debug:
        display('Insurance Flag {}: {}'.format(component, insurance_flag))

    # Insurance Rate
    try:
        insurance_rate = Inputs[
            (Inputs['Sub-system'] == subsystem) &
            (Inputs['Element'] == element) &
            (Inputs['Component'] == component) &
            (Inputs['Description'] == 'Insurance Rate')].Number.item()

    except:
        insurance_rate = 1

    if Debug:
        display('Insurance Rate {}: {}'.format(component, insurance_rate))

    # Decommissioning
    try:
        decommissioning = Inputs[
            (Inputs['Sub-system'] == subsystem) &
            (Inputs['Element'] == element) &
            (Inputs['Component'] == component) &
            (Inputs['Description'].str.contains('decommissioning'))].Number.item()

    except:
        decommissioning = 1

    if Debug:
        display('Decommissioning {}: {}'.format(component, decommissioning))

    # Residual Value
    try:
        residual_value = Inputs[
            (Inputs['Sub-system'] == subsystem) &
            (Inputs['Element'] == element) &
            (Inputs['Component'] == component) &
            (Inputs['Description'].str.contains('residual value'))].Number.item()

    except:
        residual_value = 0.01

    if Debug:
        display('Residual Value {}: {}'.format(component, residual_value))

    # -----------------------------------------------
    # generate a list of escalation factors
    previous = 1
    escalation_list = []
    escalation_years = []
    for index, year in enumerate(list(range(escalation_base_year, escalation_base_year + lifecycle + 1))):
        previous = previous * (1 + escalation_rate)
        escalation_list.append(previous)
        escalation_years.append(year)

    if Debug:
        display('Escalation years: {}'.format(escalation_years))
        display('Escalation values: {}'.format(escalation_list))

    # CAPEX per unit
    # NB: we may want to separate these later (if we want to show which components are most influential)
    try:
        Capex_per_unit = Inputs[
            (Inputs['Sub-system'] == subsystem) &
            (Inputs['Element'] == element) &
            (Inputs['Component'] == component) &
            (Inputs['Description'].str.contains(
                '|'.join(capex_categories)))].Number.sum()

    except:
        Capex_per_unit = 1_500_000 * Number_of_units

    if Debug:
        display('CAPEX total {}: {} eu per {}'.format(component, Capex_per_unit, Units))

    # initialise revenue values
    revenue_years = []
    revenue_values = []

    # generate CAPEX values
    capex_years = list(range(startyear, startyear + Construction_duration))
    capex_values = [-item * Capex_per_unit * Number_of_units for item in Construction_allocation]

    if Debug:
        display('CAPEX years: {}'.format(capex_years))
        display('CAPEX values: {}'.format(capex_values))

    # implement reinvestment here
    year = startyear
    investmentyear = startyear
    while year <= escalation_base_year + lifecycle:
        if year == escalation_base_year + lifecycle:
            # decommission
            print('decommmissioning in {}'.format(year))
            capex_years.append(year)
            capex_values.append(-Capex_per_unit * Number_of_units * decommissioning)

            revenue_years.append(year)
            revenue_values.append(Capex_per_unit * Number_of_units * ((Economic_lifetime - (year - investmentyear)) / Economic_lifetime))

            print('Residual value {}'.format(Capex_per_unit * Number_of_units * ((Economic_lifetime - (year - investmentyear)) / Economic_lifetime)))
        elif year == investmentyear + Economic_lifetime:
            # reinvest
            print('reinvest in {}'.format(year))
            capex_years.append(year)
            capex_values.append(-Capex_per_unit * Number_of_units)
            investmentyear = year

        year = year + 1

    # escalate the CAPEX using the list of escalation factors
    for i, capex_year in enumerate(capex_years):
        capex_values[i] = capex_values[i] * escalation_list[
            [index for index, escalation_year in enumerate(escalation_years) if escalation_year == capex_year][0]]

    if Debug:
        display('CAPEX years escalated: {}'.format(capex_years))
        display('CAPEX values escalated: {}'.format(capex_values))

    # use the sum of the escalated CAPEX values as OPEX value
    opex_value = sum(capex_values) * Inputs[
        (Inputs['Sub-system'] == subsystem) &
        (Inputs['Element'] == element) &
        (Inputs['Component'] == component) &
        (Inputs['Description'].str.contains(
            '|'.join(opex_categories)))].Number.sum()

    opex_years = list(range(startyear + Construction_duration, escalation_base_year + lifecycle + 1))
    opex_values = [opex_value] * len(opex_years)

    # escalate the OPEX using the list of escalation factors
    for i, opex_year in enumerate(opex_years):
        opex_values[i] = opex_values[i] * escalation_list[
            [index for index, escalation_year in enumerate(escalation_years) if escalation_year == opex_year][0]]

    if Debug:
        display('OPEX value: {}'.format(opex_value))
        display('OPEX years escalated: {}'.format(opex_years))
        display('OPEX values escalated: {}'.format(opex_values))

    # escalate the OPEX using the list of escalation factors
    for i, revenue_year in enumerate(revenue_years):
        revenue_values[i] = revenue_values[i] * escalation_list[
            [index for index, escalation_year in enumerate(escalation_years) if escalation_year == revenue_year][0]]

    df = create_cashflow_dataframe(escalation_base_year=escalation_base_year, lifecycle=lifecycle,
                                   capex={'years': capex_years,
                                          'values': capex_values},
                                   opex={'years': opex_years,
                                       'values': opex_values},
                                   revenue={'years': revenue_years,
                                       'values': revenue_values})


    return df