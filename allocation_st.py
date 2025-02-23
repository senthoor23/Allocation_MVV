import streamlit as st
import pandas as pd
import os
import win32com.client
from termcolor import colored
 
# Function to scrape data from an Excel sheet
def scrape_data_from_excel(file):
    df = pd.read_excel(file, engine='openpyxl')
    df = df[['DMX_ISSUER_ID', 'DMX_ISSUER_NAME', 'TOTAL', 'COUNTRY_DOMICILE']]
    return df
 
# Function to allocate issuers to team members
def allocate_issuers(df, team_members):
    team_totals = {member: 0 for member in team_members}
    us_counts = {member: 0 for member in team_members}
    allocation = []
    allocated_issuers = set()
 
    level_1_countries = ['AU', 'CA', 'GB', 'HK', 'IE', 'MY', 'NZ', 'SG']
    level_2_countries = ['AE', 'AR', 'AT', 'AZ', 'BE', 'BF', 'BG', 'BH', 'BM', 'BS', 'CH', 'CL', 'CO', 'CR', 'CY', 'CZ',
                         'DE', 'DK', 'EE', 'ES', 'FI', 'FO', 'FR', 'GE', 'GG', 'GI', 'GR', 'HR', 'HU', 'ID', 'IL',
                         'IM', 'IN', 'JE', 'KE', 'KW', 'KY', 'KZ', 'LI', 'LT', 'LU', 'MA', 'MC', 'MN', 'MO',
                         'MT', 'MU', 'MX', 'NG', 'NL', 'NO', 'OM', 'PA', 'PE', 'PH', 'PK', 'PL', 'PR', 'PT', 'QA', 'RO',
                         'SA', 'SE', 'SK', 'SN', 'SV', 'TG', 'TH', 'TN', 'UA', 'UY', 'VG', 'PG', 'CI']
    level_3_countries = ['BR', 'CN', 'EG', 'IT', 'RU', 'TR', 'TW', 'ZA', 'IS']
 
    us_issuers = df[df['COUNTRY_DOMICILE'] == 'US']
    level_1_issuers = df[df['COUNTRY_DOMICILE'].isin(level_1_countries)]
    level_2_issuers = df[df['COUNTRY_DOMICILE'].isin(level_2_countries)]
    level_3_issuers = df[df['COUNTRY_DOMICILE'].isin(level_3_countries)]
    other_issuers = df[~df['COUNTRY_DOMICILE'].isin(['US'] + level_1_countries + level_2_countries + level_3_countries)]
 
    us_index = 0
    while us_index < len(us_issuers):
        for member in sorted(team_members, key=lambda x: (us_counts[x], team_totals[x])):
            if us_index < len(us_issuers):
                row = us_issuers.iloc[us_index]
                allocation.append((row['DMX_ISSUER_ID'], row['DMX_ISSUER_NAME'], row['TOTAL'], row['COUNTRY_DOMICILE'], member))
                team_totals[member] += row['TOTAL']
                us_counts[member] += 1
                allocated_issuers.add(row['DMX_ISSUER_ID'])
                us_index += 1
 
    def allocate_by_level(issuers):
        for index, row in issuers.iterrows():
            if row['DMX_ISSUER_ID'] not in allocated_issuers:
                min_member = min(team_totals, key=team_totals.get)
                allocation.append((row['DMX_ISSUER_ID'], row['DMX_ISSUER_NAME'], row['TOTAL'], row['COUNTRY_DOMICILE'], min_member))
                team_totals[min_member] += row['TOTAL']
                allocated_issuers.add(row['DMX_ISSUER_ID'])
 
    allocate_by_level(level_1_issuers)
    allocate_by_level(level_2_issuers)
    allocate_by_level(level_3_issuers)
    allocate_by_level(other_issuers)
 
    allocation_df = pd.DataFrame(allocation, columns=['DMX_ISSUER_ID', 'DMX_ISSUER_NAME', 'TOTAL', 'COUNTRY_DOMICILE', 'Team_Member'])
    return allocation_df
 
# Function to validate the allocation
def validate_allocation(allocation_df, team_members):
    total_points = allocation_df['TOTAL'].sum()
    average_points_per_member = total_points / len(team_members)
    validation_results = {}
 
    for member in team_members:
        member_total = allocation_df[allocation_df['Team_Member'] == member]['TOTAL'].sum()
        difference_from_average = member_total - average_points_per_member
        validation_results[member] = {
            'Total': member_total,
            'Difference from Average': difference_from_average,
            'Above Average': difference_from_average > 0,
            'Below Average': difference_from_average < 0
        }
    return validation_results, average_points_per_member
 
# Streamlit Interface
st.title("Issuer Allocation System")
 
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])
team_members = st.text_input("Enter Team Members (comma-separated):").split(',')
 
if uploaded_file is not None and team_members:
    df = scrape_data_from_excel(uploaded_file)
    allocation_df = allocate_issuers(df, team_members)
    allocation_df = allocation_df.set_index('DMX_ISSUER_ID').reindex(df['DMX_ISSUER_ID']).reset_index()
    validation_results, avg_points = validate_allocation(allocation_df, team_members)
 
    st.subheader("Allocation Results")
    st.dataframe(allocation_df)
 
    st.subheader("Validation Results")
    st.write(f"**Average Points per Member:** {avg_points:.2f}")
    for member, result in validation_results.items():
        status = "Above Average" if result['Above Average'] else "Below Average"
        color = "green" if result['Above Average'] else "red"
        st.markdown(f"**{member}**: Total - {result['Total']}, Difference from Average - :{color}[{result['Difference from Average']:.2f}] ({status})")
 
    st.download_button(
        label="Download Allocation Result",
        data=allocation_df.to_csv(index=False).encode('utf-8'),
        file_name='allocation_results.csv',
        mime='text/csv'
    )