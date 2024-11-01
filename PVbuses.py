import pandas as pd
import matplotlib.pyplot as plt

# Load the Excel file
file_path = 'C:\\Apps-SU\\My workspace\\res_sgen\\p_mw.xlsx'
data = pd.read_excel(file_path, usecols="B:F")


# Define PV buses and titles for the plot
pv_buses = [1, 21, 26, 22, 12]

# Plotting each column as a line plot for each bus
# plt.figure(figsize=(10, 6))
# for i, bus in enumerate(pv_buses):
#     plt.plot(data.index, data.iloc[:, i], label=f'Bus {bus}')
    
    
# Plotting the first 240 data points for each bus
plt.figure(figsize=(10, 6))
for i, bus in enumerate(pv_buses):
    plt.plot(data.index[:240], data.iloc[:240, i], label=f'Bus {bus}')

# Setting labels and title
plt.xlabel('Hours (2023)')
plt.ylabel('PV Generation (MW)')
plt.title('Hourly PV Generation (MW) for Selected Buses (8760 Data Points)')
plt.legend(title="PV Buses")
plt.grid(True)
plt.show()
