import os
import pandas as pd
import pandapower as pp
import pandapower.networks as pn
import pandapower.plotting as plot
from pandapower.control import ConstControl
from pandapower.timeseries import DFData, OutputWriter, run_timeseries
import matplotlib.pyplot as plt
import pvlib
# ========== Power network =============

# Create PandaPower network (IEEE-30bus)
net = pp.networks.case30()
# View all generator (gen) information
gen_info = net.gen
# Print gen information
print(gen_info)

# Print the bus number where the PV generator is located
pv_bus_numbers = net.gen['bus'].values
print("Generator (PV) bus numbers:", pv_bus_numbers)
# Identify the 5 PV buses in the IEEE 30-bus system
pv_buses = [1, 21, 26, 22, 12]

# View the bus information of the external power grid (Slack bus)
slack_bus_numbers = net.ext_grid['bus'].values
print("Slack bus numbers:", slack_bus_numbers)

# View load information
print(net.load)

# Calculate total active power and reactive power
total_p = net.load['p_mw'].sum() # Total active power
total_q = net.load['q_mvar'].sum() # Total reactive power

print(f"Total active power load of IEEE 30-bus system: {total_p} MW")
print(f"Total reactive power load of IEEE 30-bus system: {total_q} MVar")

# Initialize an empty list to store installed capacities of each generator
installed_capacities = []

# Loop through each row in the gen_info DataFrame to extract information
for i, row in gen_info.iterrows():
    bus_id = row['bus']  # Get the bus ID for the generator
    installed_capacity = row['max_p_mw']  # Get the installed capacity of the generator in MW
    print(f"Generator at Bus {bus_id}: Installed Capacity = {installed_capacity} MW")
    installed_capacities.append(installed_capacity)  # Append the capacity to the list

# Calculate the total installed capacity of all generators
total_installed_capacity = sum(installed_capacities)
print(f"Total Installed Capacity = {total_installed_capacity} MW")

# ========== Load data and data processing =============

# Load the CSV file with ERCOT time-series data file
file_path = 'C:\\Users\\jliu359\\OneDrive - Syracuse University\\Desktop\\ERCOT SOLAR DATA.xlsx'
ercot_data == pd.read_excel(file_path)
#  Read data
print(ercot_data.head())

# Scale ERCOT's generation to match the total installed capacity of the 5 PV buses
scaling_factor = total_installed_capacity / 14249  # 14249 MW is the total solar installed capacity in the ERCOT dataset
scaled_pv_generation = pv_generation * scaling_factor

# Extract the time and PV generation columns
time_series = ercot_data['Time (Hour-Ending)']
pv_generation = ercot_data['ERCOT.PVGR.GEN']


# selected_columns_corrected = ['Dew Point', 'GHI', 'DHI', 'DNI', 'Pressure', 'Temperature']
# nsrdb_data = data[selected_columns_corrected]
#
# # Confirm that data
# print(nsrdb_data.head())
# ghi = nsrdb_data['GHI']
# dhi = nsrdb_data['DHI']
# dni = nsrdb_data['DNI']
# temperature = nsrdb_data['Temperature']
# pressure = nsrdb_data['Pressure']
# dew_point = nsrdb_data['Dew Point']
#
# # Ensure datas are numeric
# # Convert all relevant columns to numeric, coercing errors
#
# ghi = pd.to_numeric(nsrdb_data['GHI'], errors='coerce')
# dhi = pd.to_numeric(nsrdb_data['DHI'], errors='coerce')
# dni = pd.to_numeric(nsrdb_data['DNI'], errors='coerce')
# temperature = pd.to_numeric(nsrdb_data['Temperature'], errors='coerce')
# pressure = pd.to_numeric(nsrdb_data['Pressure'], errors='coerce')
# dew_point = pd.to_numeric(nsrdb_data['Dew Point'], errors='coerce')
#
# # Check for NaN values
# print(ghi.isna().sum(), "NaN values in GHI")
# print(dhi.isna().sum(), "NaN values in DHI")
# print(dni.isna().sum(), "NaN values in DNI")
#
# # Define the PV system
# latitude = 43.04 # Latitude of Syracuse
# longitude = -76.15 # Longitude of Syracuse
# location = pvlib.location.Location(latitude, longitude)
#
# # Create a PV system and calculate photovoltaic power generation based on meteorological data
# system = pvlib.pvsystem.PVSystem(surface_tilt=43, surface_azimuth=180)
#
# # Time series
# times = pd.date_range('2022-01-01', periods=len(nsrdb_data ), freq='h')
#
# # Get the sun position
# solar_position = location.get_solarposition(times)
#
# # Use meteorological data to calculate the power generation of the photovoltaic system
# poa = system.get_irradiance(dni=dni, ghi=ghi, dhi=dhi, solar_zenith=solar_position['zenith'], solar_azimuth=solar_position['azimuth'])
#
# # Assume cell temperature as ambient temperature for simplicity (you can model this more precisely if needed)
# temp_cell = temperature
#
# # Define system parameters for DC and AC power calculations
# pdc0 = 4000  # DC system size (Watts)
# gamma_pdc = -0.004  # Power temperature coefficient (%/°C)
#
# # Calculate DC power using pvwatts_dc
# dc_power = pvlib.pvsystem.pvwatts_dc(g_poa_effective=poa['poa_global'], temp_cell=temp_cell, pdc0=pdc0, gamma_pdc=gamma_pdc)
#
# # Define inverter efficiency parameters
# eta_inv_nom = 0.96  # Nominal inverter efficiency
# eta_inv_ref = 0.9637  # Reference inverter efficiency
#
# # Calculate AC power using pvwatts_ac, converting DC power to AC
# ac_power =dc_power*eta_inv_nom
#
# # Print power generation results
# print(ac_power)
#
# # Run OPF
# pp.runopp(net)
#
# # Check results
# print(net.res_bus)  # Voltage results at each bus
# print(net.res_gen)  # Generation results at each bus
# print(net.res_sgen)  # PV bus results