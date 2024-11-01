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
total_p = net.load['p_mw'].sum()  # Total active power
total_q = net.load['q_mvar'].sum()  # Total reactive power

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

# # Remove existing generators on PV buses to avoid duplicates
# for bus in pv_buses:
#     net.gen.drop(net.gen[net.gen['bus'] == bus].index, inplace=True)

# Create  Generators (gen) for the PV buses（define the Voltage Set-point）
for bus, installed_capacity in zip(pv_buses, installed_capacities):
    pp.create_gen(net, bus, p_mw=0, vm_pu=1.00, max_p_mw=installed_capacity, name=f"PV_bus_{bus}")

# View all gen information
gen_info = net.gen
print(gen_info)

# ========== Load data and data processing =============

# Load the CSV file with ERCOT time-series data file
# file_path = 'C:\\Users\\default.DESKTOP-C4C7JDR\\Desktop\\ERCOT SOLAR DATA.xlsx'  # Home PC
file_path = 'C:\\Users\\jliu359\\OneDrive - Syracuse University\\Desktop\\ERCOT SOLAR DATA.xlsx' #Lab PC
ercot_data = pd.read_excel(file_path)
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
        element_index = net.gen[net.gen['bus'] == bus].index[0]
        column_name = f'bus_{bus}_p_mw'
        ConstControl(net, element='gen', variable='p_mw', element_index=element_index, data_source=data_source,
                     profile_name=column_name)

# Set up the output writer for results
# output_dir = "D:/My workspace/Results"#Home PC
output_dir = "C:/Apps-SU/My workspace/" #Lab PC
ow = OutputWriter(net, output_path=output_dir, output_file_type=".xlsx")
# Define Record
ow.log_variable('res_gen', 'p_mw')

# Define the number of time steps for the simulation
# time_steps = range(len(profiles_df))  # Define time steps as a range matching the data length
time_steps = list(range(0, 8760))
# Run the time-series simulation with the defined number of time steps
run_timeseries(net, time_steps=time_steps)

# ========== Plot the 5PV buses power function =============
# Define the output directory and updated file path for res_gen
output_dir = "C:/Apps-SU/My workspace/res_gen/"
result_file_path = os.path.join(output_dir, "p_mw.xlsx")
# Assuming the Excel file has columns corresponding to the first 5 buses for PV generation profiles
try:
    # Load the data from p_mw.xlsx for plotting
    pv_generation_data = pd.read_excel(result_file_path, usecols="B:F")  # Load columns B to F (2nd to 6th columns)
    plt.figure()
    plt.title('PV Generation Profile at Each PV Bus')

    # Plot each bus's generation profile from the loaded data
    for i, bus in enumerate(pv_buses):
        plt.plot(pv_generation_data.iloc[:, i], label=f'Bus {bus}')  # Use column for buses

    plt.xlabel('Time')
    plt.ylabel('PV Generation (MW)')
    plt.legend()
    plt.show()

except FileNotFoundError:
    print(f"File not found: {result_file_path}")

# ========== OPF with Voltage Constraints =============

# Set voltage constraints to ±5% around 1.00 p.u.
# net.bus['min_vm_pu'] = 0.95
# net.bus['max_vm_pu'] = 1.05

# Perform optimal power flow (OPF) analysis
pp.runopp(net)

# Display results from the OPF
if not pp.runopp(net):
    print("OPF did not converge.")
else:
    print("OPF Results:")
    print(net.res_bus)
    print(net.res_line)
