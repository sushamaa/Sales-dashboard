import streamlit as st
import plotly.express as px
import pandas as pd
import openpyxl 
import functools
from st_aggrid import AgGrid
from st_aggrid.shared import JsCode
from st_aggrid.grid_options_builder import GridOptionsBuilder
import warnings

# Suppress future warnings for clean output
warnings.filterwarnings("ignore", category=FutureWarning, module="pyarrow")

# Configure the Streamlit app
st.set_page_config(page_title="Performance Dashboard", page_icon=":bar_chart:", layout="wide")

# Load custom CSS
with open('style.css') as f:
  st.markdown(f'<style>{f.read()}</style>', unsafe_allow_html=True)

# Main page title and separator
st.markdown("### :bar_chart: Performance Dashboard")
st.markdown("---")

# Function to get data from Excel file
@st.cache_data
def get_data_from_excel():
  df = pd.read_excel(io ='Book1.xlsx',sheet_name='performance', engine='openpyxl', usecols='A:H', nrows=20627)
  return df

# Load data from Excel files
df = get_data_from_excel()
df1 = pd.read_excel(io = 'Book1.xlsx',sheet_name='target', engine = 'openpyxl', usecols='A:E', nrows=76)

# Exclude specific branches
df = df[(df['Branch']!='COP') & (df['Branch']!='VTZ')]

# Function to map branches to regions
region_branch = {'North':['AMD','BRC','DEL','GDM','JAI'],'East':['KOL I','KOL II'],'West':['HYD','MUM','PNQ'],'South':['BLR','CJB','COK','MAA','TUT']}
def regions(value):
    for j in region_branch:
      if value in region_branch[j]:
        return j

# Map branches to their regions 
df.loc[:, 'Region'] = df['Branch'].map(regions)

# Reorder columns
df = df[['Region', 'Branch', 'Routed_By','Mode','to_be_consider','Month','Volume','Net_Profit','Recent_Month']]

# Sidebar filter
st.sidebar.header("Please Filter Here: ")
## "Please note that altering the mode will result in modifications to the entire report."
a1_export = sorted(df['Mode'].unique())
export = st.sidebar.selectbox("Select Type of Export:", options = a1_export)

a2_month = sorted(df['Month'].unique())
month = st.sidebar.multiselect("Choose Month:", options = a2_month, default = a2_month[-3:])

a3_route = sorted(df['Routed_By'].unique())
route = st.sidebar.multiselect("Choose Route:", options = a3_route, default = a3_route)

a4_region = df['Region'].unique()
region = st.sidebar.multiselect("Choose Region:", options = a4_region, default = a4_region[0:1])

# Filter data based on sidebar selections
df_selection = df.query("Mode == @export & Month == @month & Routed_By == @route & Region == @region")

# Function to assign unit value to volume based on export type
def unit(export):
  if export[0:3]=='FCL':
    return "TEU's"
  if export[0:3]=='LCL':
    return 'CBM'
  else:
    return "Tonnes"

# Sidebar KPI's
total_revenue = round(df_selection.loc[df_selection['to_be_consider'] == 'Y', 'Net_Profit'].sum()/100000,2)
total_volume = round(df_selection['Volume'].sum())
average_revenue = round((total_revenue/len(month)),2)
average_volume = round(total_volume/len(month))

st.sidebar.markdown("## Summary of selected data:")

st.sidebar.markdown(f"### Total Revenue: {total_revenue:,} lakhs")

st.sidebar.markdown(f"### Total Volume: {total_volume:,} {unit(export)}")

st.sidebar.markdown(f"### Average Revenue: {average_revenue:,} lakhs")

st.sidebar.markdown(f"### Average Volume: {average_volume:,} {unit(export)}")  

# Functions to calculate targets
def regional_revenue_target(region1, mode1):
  total = round(df1.loc[(df1['REGION'] == region1) & (df1['EXPORT'] == mode1) , 'REVENUE'].sum(),2)
  return total

def regional_volume_target(region1, mode1):
  total = round(df1.loc[(df1['REGION'] == region1) & (df1['EXPORT'] == mode1) , 'VOLUME1'].sum())
  return total

def branch_revenue_target(region1, mode1, branch1):
  total = round(df1.loc[(df1['REGION'] == region1) & (df1['EXPORT'] == mode1) & (df1['BRANCH'] == branch1) , 'REVENUE'].sum(),2)
  return total

def branch_volume_target(region1, mode1, branch1):
  total = round(df1.loc[(df1['REGION'] == region1) & (df1['EXPORT'] == mode1) & (df1['BRANCH'] == branch1) , 'VOLUME1'].sum())
  return total

# Achieved vs Target selection
st.markdown(f"### Achieved vs Target%")
s1 = st.selectbox("Peformance of:", options = [(f"Last Month ({a2_month[-2]})"), (f"Current Month ({a2_month[-1]}) (1ST TO 10TH)")])

# Top KPIs for revenue and volume
regions = ['North','South','East','West']
col1, col2, col3, col4 = st.columns(4)

# Calculate and display revenue KPIs
if s1 == (f"Last Month ({a2_month[-2]})"):
  for col, region in zip([col1, col2, col3, col4], regions):
    with col:
      rev = round(df.loc[(df['Region'] == region) & (df['Mode'] == export) & (df['Month'] == a2_month[-2]) & (df['to_be_consider'] == 'Y'), 'Net_Profit'].sum()/100000,2)
      try:
        final = round((rev/regional_revenue_target(region,export))*100)
        if final >= 100:
          st.metric(label = (f"{region} Revenue"), value = (f"{rev} lakhs") , delta = (f"{final}%"))
        else:
          st.metric(label = (f"{region} Revenue"), value = (f"{rev} lakhs") , delta = (f"{final}%"), delta_color='inverse')
      except:
        st.metric(label = (f"{region} Revenue"), value = (f"{rev} lakhs"))
      st.markdown(f"##### Target: {round(df1.loc[(df1['REGION'] == region) & (df1['EXPORT'] == export), 'REVENUE'].sum())} lakhs")
else:
  for col, region in zip([col1, col2, col3, col4], regions):
    with col:
      rev = round(df.loc[(df['Region'] == region) & (df['Mode'] == export) & (df['Month'] == a2_month[-1]) & (df['to_be_consider'] == 'Y'), 'Net_Profit'].sum()/100000,2)
      try:
        final = round((rev/regional_revenue_target(region,export))*100)
        if final >= 100:
          st.metric(label = f"{region} Revenue", value = (f"{rev} lakhs") , delta = (f"{final}%"))
        else:
          st.metric(label = f"{region} Revenue", value = (f"{rev} lakhs") , delta = (f"{final}%"),  delta_color='inverse')
      except:
        st.metric(label = f"{region} Revenue", value = (f"{rev} lakhs"))
      st.markdown(f"##### Target: {round(df1.loc[(df1['REGION'] == region) & (df1['EXPORT'] == export), 'REVENUE'].sum())} lakhs")

# Calculate and display volume KPIs
col11, col22, col33, col44 = st.columns(4)
if s1 == (f"Last Month ({a2_month[-2]})"):
  for col, region in zip([col11, col22, col33, col44], regions):
    with col:
      vol = round(df.loc[(df['Region'] == region) & (df['Mode'] == export) & (df['Month'] == a2_month[-2]), 'Volume'].sum())
      try:
        final = round((vol/regional_volume_target(region, export))*100)
        if final >= 100:
          st.metric(label = f"{region} Volume", value = (f"{vol} {unit(export)}") , delta = (f"{final}%"))
        else:
          st.metric(label = f"{region} Volume", value = (f"{vol} {unit(export)}") , delta = (f"{final}%"),  delta_color='inverse')
      except:
        st.metric(label = f"{region} Volume", value = (f"{vol} {unit(export)}"))
      st.markdown(f"##### Target: {round(df1.loc[(df1['REGION'] == region) & (df1['EXPORT'] == export), 'VOLUME1'].sum())} {unit(export)}")
else:
  for col, region in zip([col11, col22, col33, col44], regions):
    with col:
      vol = round(df.loc[(df['Region'] == region) & (df['Mode'] == export) & (df['Month'] == a2_month[-1]) & (df['Recent_Month'] == '1ST TO 10TH'), 'Volume'].sum())
      try:
        final = round((vol/regional_volume_target(region, export))*100)
        if final >= 100:
          st.metric(label = f"{region} Volume", value = (f"{vol} {unit(export)}") , delta = (f"{final}%"))
        else:
          st.metric(label = f"{region} Volume", value = (f"{vol} {unit(export)}") , delta = (f"{final}%"),  delta_color='inverse')
      except:
        st.metric(label = f"{region} Volume", value = (f"{vol} {unit(export)}"))
      st.markdown(f"##### Target: {round(df1.loc[(df1['REGION'] == region) & (df1['EXPORT'] == export), 'VOLUME1'].sum())} {unit(export)}")

# Add a horizontal line separator
st.markdown("---")  

st.markdown('###### To see the complete table, simply click on the drop-down menu.')

# Function to generate table for specific region and export mode
def table1(region_1,export_1):
  branch_of_that_region = sorted(df1.loc[(df1['REGION'] == region_1) & (df1['EXPORT'] == export_1) , 'BRANCH'].unique())
  current_rev,current_vol,last_rev,last_vol,target_rev,target_vol = [],[],[],[],[],[]
  ach_rev_current, ach_vol_current,ach_rev_last, ach_vol_last  = [], [], [], []
  for i in branch_of_that_region:
    current_rev.append(round(df.loc[(df['Branch'] == i) & (df['Mode'] == export_1) & (df['to_be_consider'] =='Y') & (df['Month'] == a2_month[-1]), 'Net_Profit'].sum()/100000,2))
    current_vol.append(round(df.loc[(df['Branch'] == i) & (df['Mode'] == export_1) & (df['Recent_Month'] =='1ST TO 10TH') & (df['Month'] == a2_month[-1]), 'Volume'].sum()))
    last_rev.append(round(df.loc[(df['Branch'] == i) & (df['Mode'] == export_1) & (df['to_be_consider'] =='Y') & (df['Month'] == a2_month[-2]), 'Net_Profit'].sum()/100000,2))
    last_vol.append(round(df.loc[(df['Branch'] == i) & (df['Mode'] == export_1) & (df['Month'] == a2_month[-2]), 'Volume'].sum()))
    target_rev.append(round(df1.loc[(df1['BRANCH'] == i) & (df1['EXPORT'] == export_1), 'REVENUE'].sum(),2))
    target_vol.append(round(df1.loc[(df1['BRANCH'] == i) & (df1['EXPORT'] == export_1), 'VOLUME1'].sum()))
  for i in range(len(branch_of_that_region)):
    if (target_rev[i]!=0):
      ach_rev_current.append(round((current_rev[i]/target_rev[i])*100))
      ach_rev_last.append(round((last_rev[i]/target_rev[i])*100))
    else:
      ach_rev_current.append(0)
      ach_rev_last.append(0)
    if (target_vol[i]!=0):
      ach_vol_current.append(round((current_vol[i]/target_vol[i])*100))
      ach_vol_last.append(round((last_vol[i]/target_vol[i])*100))
    else:
      ach_vol_current.append(0)
      ach_vol_last.append(0)

   # Create a DataFrame with calculated values
  d = {'Branch':branch_of_that_region,'Target_revenue':target_rev,'Target_volume':target_vol,(f'Rev%({a2_month[-1]})'): ach_rev_current,(f'Vol%({a2_month[-1]})'):ach_vol_current,
  (f'Rev%({a2_month[-2]})'):ach_rev_last,(f'Vol%({a2_month[-2]})'):ach_vol_last,(f'Revenue({a2_month[-1]})'):current_rev,(f'Volume({a2_month[-1]})'):current_vol,(f'Revenue({a2_month[-2]})'):last_rev,(f'Volume({a2_month[-2]})'):last_vol,}
  return pd.DataFrame(data=d)

# Iterate over each region and create an expander for performance details
for i in ['North','South','East','West']:
  with st.expander(f"{i} Performance"):
    st.subheader(f"{i} Region:")

# Define JavaScript code for cell styling in the grid
cellsytle_jscode = JsCode("""
    function(params) {
        const cellStyle = {
            'color': 'white',
            'display': 'flex',
            'alignItems': 'center',
            'justifyContent': 'center',
            'padding': '0',
            'margin': '0'
        };

        if (params.value > 99) {
            cellStyle.backgroundColor = 'forestgreen';
        } else if (params.value < 100) {
            cellStyle.backgroundColor = 'crimson';
        } else {
            cellStyle.backgroundColor = 'slategray';
        }

        return cellStyle;
    }
""")

    # Build the grid options for the data table
    gb = GridOptionsBuilder.from_dataframe(table1('North',export))
    gb.configure_columns(((f'Rev%({a2_month[-1]})'),(f'Vol%({a2_month[-1]})'),(f'Rev%({a2_month[-2]})'),(f'Vol%({a2_month[-2]})')),cellStyle=cellsytle_jscode,)
    gb.configure_pagination()
    gb.configure_columns(("Branch",'Target_revenue','Target_volume'), pinned=True)
    gridOptions = gb.build()

    # Display the data table using AgGrid
    AgGrid(table1(i,export), gridOptions=gridOptions, allow_unsafe_jscode=True)
    st.markdown("** Revenue in lakhs")

st.markdown("---")

# Section for performance of Agent and Swift
st.markdown('#### Performance of Agent and Swift:')
col1,col2 = st.columns(2)

# Dropdown to select the region
with col1:
  the_region = st.selectbox("Select Region:", options = a4_region)

# Dropdown to select the month
with col2:
  the_month = st.selectbox("Select Month:", options = a2_month)

# Filter data based on selected month and region
new_df = df.query("Month == @the_month & Region == @the_region")  
new_df.loc[:, 'Net_Profit'] = new_df['Net_Profit'].div(100000).round(2)

# Create charts for the selected data
chart = functools.partial(st.plotly_chart, use_container_width=True)
def graph1(the_month, export, the_region):
  col1a, col2a = st.columns(2)
    ## for revenue
  with col1a:
    st.markdown('#### Performance in terms of revenue:')
    data = new_df[(new_df['Month'] == the_month) & (new_df['Region'] == the_region) & (new_df['Mode'] == export) & (new_df['to_be_consider'] == 'Y')]
    fig = px.sunburst(data, path=['Branch', 'Routed_By'], values='Net_Profit', color='Branch')
    chart(fig)  
    st.table(data.groupby(['Branch', 'Routed_By'])['Net_Profit'].sum().reset_index())
    st.markdown("** Revenue in lakhs")

    ## for volume
  with col2a:
    if (the_month == a2_month[-1]):
      st.markdown('#### Performance in terms of Volume:')
      data = new_df[(new_df['Month'] == the_month) & (new_df['Region'] == the_region) & (new_df['Mode'] == export) & (new_df['to_be_consider'] == 'Y') & (new_df['Recent_Month'] == '1ST TO 10TH')]
      fig = px.sunburst(data, path=['Branch', 'Routed_By'], values='Volume', color='Branch')
      chart(fig)
      st.table(data.groupby(['Branch', 'Routed_By'])['Volume'].sum().reset_index())
    else:
      st.markdown('#### Performance in terms of Volume:')
      data = new_df[(new_df['Month'] == the_month) & (new_df['Region'] == the_region) & (new_df['Mode'] == export) & (new_df['to_be_consider'] == 'Y')]
      fig = px.sunburst(data, path=['Branch', 'Routed_By'], values='Volume', color='Branch')
      chart(fig)
      st.table(data.groupby(['Branch', 'Routed_By'])['Volume'].sum().reset_index())

# Call the function to create graphs
graph1(the_month, export, the_region)

# Add a disclaimer
st.markdown("---")
st.markdown("#### Disclaimer: The data presented in this document or any associated content is entirely fictional and for illustrative purposes only. Any resemblance to real individuals, organizations, or events is purely coincidental. This data should not be used for any real-world decisions, analysis, or actions, as it is entirely fabricated and does not reflect actual information or circumstances.")
st.markdown("---")
