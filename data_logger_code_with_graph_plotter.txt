import pandas as pd
import openpyxl
import matplotlib.pyplot as plt

#################################################################################################################################

df = pd.read_excel(r'C:\Users\au000tn4\OneDrive - WSA\Intern\Test and experiments\MEMSDATAlowhz.xlsx')
wrkbk = openpyxl.load_workbook(r'C:\Users\au000tn4\OneDrive - WSA\Intern\Test and experiments\MEMSDATAlowhz.xlsx')

#remember to check filename(MEMSDATA) and location, DELETE FIRST ROW, save and close Excel file


#######################################################  INPUTS  ###############################################################
seconds_to_search = 1             
recording_rate_per_second = 200    #Microphone measurement interval 
accepted_diff = 150 
range_of_graph_in_seconds = 1  #Number of seconds shown before and after data point of interest on graph
minutes_to_skip_at_start = 0    #Number of minutes to skip from start of test (nearest whole number)
testend_datapoints_to_remove = 0   #No. of data points at end to not be plotted. Number = seconds * recording rate

#minutes_to_remove_at_end = 0 
#testend_datapoints_to_remove = minutes_to_remove_at_end * recording_rate_per_second * 60
################################################################################################################################

dataEx = wrkbk.active
data = []
diff_list = []
difflistindex = []
n = int(seconds_to_search * recording_rate_per_second)
s = range_of_graph_in_seconds * recording_rate_per_second
startpoint = minutes_to_skip_at_start * recording_rate_per_second * 60

for i in range(1, dataEx.max_row+1):     #Displays raw data from excel in original form
    cell_obj = dataEx.cell(row=i, column=2)
    data.append(cell_obj.value)
        
for i in range(startpoint, len(data)):
    diff = max(data[i:i+n]) - min(data[i:i+n])
    if diff < accepted_diff:
        difflistindex.append(i+1)     #making list of index positions starting 1, where the set of n has a smaller difference
    
print("Note: Multiply values by 0.25 to get seconds as we measure every quarter second. Index location of max-min < " + str(accepted_diff) + " is: " + "\n")
#print(str(difflistindex) + '\n')

for i in range(0, testend_datapoints_to_remove):
    difflistindex.pop()
print('Timestamps after popping end datapoints is ' + str(difflistindex) + '\n')

for i in difflistindex:
    print('Timestamp is ' + str(i))
    plotpointslist = []
    time = []
    if i < s:
        for j1 in range(0, i+s+1):
            #print('time is ' + str(j))
            time.append(j1)
            #print('data[j] is ' + str(data[j]))
            plotpointslist.append(data[j1])
    else:
        for j in range(i-s, i+s+1):      #number of data points before and after point of interest to view in graph     
            #print('time is ' + str(j))
            time.append(j)
            #print('data[j] is ' + str(data[j]))
            plotpointslist.append(data[j])
    #print('time is ' + str(time))
    #print('plotpointslist is ' + str(plotpointslist))
    plt.plot(time, plotpointslist, color='blue')
    plt.title('Audio Value vs Time', fontsize=14)
    plt.xlabel('Time', fontsize=14)
    plt.ylabel('Audio Value', fontsize=14)
    plt.grid(True)
    plt.show()