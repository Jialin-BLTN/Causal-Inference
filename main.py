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

# Create PandaPower network (30bus)
net = pn.case30()
# View all generator (gen) information
gen_info = net.gen
print(gen_info)

# Print the bus number where the generator is located
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

#Create Static Generators (sgen) for the PV buses
for bus, installed_capacity in zip(pv_buses, installed_capacities):
    pp.create_sgen(net, bus, p_mw=0, max_p_mw=installed_capacity, name=f"PV_bus_{bus}")

# View all sgen information
sgen_info = net.sgen
print(sgen_info)

# ========== Load data and data processing =============

# Load the CSV file with ERCOT time-series data file
file_path = 'C:\\Users\\default.DESKTOP-C4C7JDR\\Desktop\\ERCOT SOLAR DATA.xlsx' #Home PC
# file_path = 'C:\\Users\\jliu359\\OneDrive - Syracuse University\\Desktop\\ERCOT SOLAR DATA.xlsx' #Lab PC
ercot_data =pd.read_excel(file_path)
#  Read data
print(ercot_data.head())

# Extract the time and PV generation columns
# Set the time column as the index and ensure it's in datetime format
ercot_data['Time (Hour-Ending)'] = pd.to_datetime(ercot_data['Time (Hour-Ending)'])
ercot_data.set_index('Time (Hour-Ending)', inplace=True)
pv_generation = ercot_data['ERCOT.PVGR.GEN']

# Scale ERCOT's generation to match the total installed capacity of the 5 PV buses
scaling_factor = total_installed_capacity / 14249  # 14249 MW is the total solar installed capacity in the ERCOT dataset
scaled_pv_generation = pv_generation * scaling_factor

# Prepare a DataFrame to hold profiles for each PV bus with an integer index
profiles_df = pd.DataFrame(index=range(len(ercot_data)))

# Assign the scaled generation to the 5 PV buses based on their individual capacities
for i, bus in enumerate(pv_buses):
    # Calculate the proportion of generation for each PV bus
    bus_scaling_factor = installed_capacities[i] / total_installed_capacity
    bus_generation_profile = scaled_pv_generation * bus_scaling_factor

    # Add the generation profile to the DataFrame with a consistent column name
    column_name = f'bus_{bus}_p_mw'
    profiles_df[column_name] = bus_generation_profile.values

    # Create a DFData object with the profiles DataFrame
    data_source = DFData(profiles_df)

    # Create ConstControl objects for each PV bus to update active power over time
    for i, bus in enumerate(pv_buses):
        element_index = net.sgen[net.sgen['bus'] == bus].index[0]
        column_name = f'bus_{bus}_p_mw'
        ConstControl(net, element='sgen', variable='p_mw', element_index=element_index, data_source=data_source,
                     profile_name=column_name)

# Set up the output writer for results
output_dir = "D:/My workspace/Results"
ow = OutputWriter(net, output_path=output_dir, output_file_type=".xlsx")
# Define record
ow.log_variable('res_sgen', 'p_mw')

# Define the number of time steps for the simulation
# time_steps = range(len(profiles_df))  # Define time steps as a range matching the data length
time_steps = list(range(0, 8760))
# Run the time-series simulation with the defined number of time steps
run_timeseries(net, time_steps=time_steps)

# Plotting the results
plt.figure()
plt.title('PV Generation Profile at Each PV Bus')
# Iterate through each PV bus to plot its time-series data
for i, bus in enumerate(pv_buses):
    element_index = net.sgen[net.sgen['bus'] == bus].index[0]
    result_file = os.path.join(output_dir, f"res_sgen_{element_index}.xlsx")
    sheet_name = f"sgen_p_mw_{element_index}"
    
    try:
        # Read the generated Excel file for the specific PV bus
        plot_timeseries = pd.read_excel(result_file)
        plt.plot(ercot_data.index, plot_timeseries['p_mw'], label=f'Bus {bus}')
    except FileNotFoundError:
        print(f"File or sheet not found: {result_file}")

plt.xlabel('Time')
plt.ylabel('PV Generation (MW)')
plt.legend()
plt.show()

# Perform optimal power flow (OPF) analysis
pp.runopp(net)

# Display results from the OPF
print("OPF Results:")
print(net.res_bus)
print(net.res_line)

