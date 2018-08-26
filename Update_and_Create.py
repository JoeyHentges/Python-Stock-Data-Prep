#this program will take an excel (csv) file and add/remove some cells

#the cells to be added are:
#day (number ex 1 2 3...)
#last_open price (higher/lower)
#next_open price (higher/lower)
#last_high price (higher/lower)
#next_high price (higher/lower)
#last_low price (higher/lower)
#next_low price (higher/lower)
#last_volume (higher/lower)
#next_volume (higher/lower)

#the cells to be removed are:
#the second row (all removed)
#the last row (all removed)

#imports
import csv

#the name of the csv file to be edited
csv_file = 'ACN_data.csv'
#the sheet number in the excel file (should be 1)


#create variables that will be added set them as empty
day = []
last_open = []
next_open = []
last_high = []
next_high = []
last_low = []
next_low = []
last_close = []
next_close = []
last_volume = []
next_volume = []

#arrays to hold excel numbers
open_array = []
high_array = []
low_array = []
close_array = []
volume_array = []

#number to hold the row number the while loop is on
row_number = 0

#open the file and count the total number of rows
total_rows = 0
with open(csv_file, 'r') as f:
    reader = csv.reader(f)
    for row in reader:
        total_rows += 1
f.close()

#run number 1
#fill the excel data holding arrays
with open(csv_file, 'r') as file:
   reader = csv.reader(file)
   for row in reader:      
        if (row_number != 0):
            #add each rows values to their corresponding arrays
            day.extend([row_number])
            open_array.extend([row[1]])
            high_array.extend([row[2]])
            low_array.extend([row[3]])
            close_array.extend([row[4]])
            volume_array.extend([row[5]])  
        row_number += 1 #update the row number counter
day.extend([row_number]) #add the last day number to the array
file.close()

#run number 2
#fill the last_open/next_open etc.. arrays
for i in range(len(day)-2):
    #check last --- and next --- for higher/lower values
    #skip first values and last values
    if(i != 0):
        if(i != len(day)-1):
            #last open
            if(open_array[i-1] >= open_array[i]):
                last_open.extend(["higher"])
            else:
                last_open.extend(["lower"])
            #-----
            #next open
            if(open_array[i+1] >= open_array[i]):
                next_open.extend(["higher"])
            else:
                next_open.extend(["lower"])
            #-----
            #last high
            if(high_array[i-1] >= high_array[i]):
                last_high.extend(["higher"])
            else:
                last_high.extend(["lower"])
            #-----
            #next high
            if(high_array[i+1] >= high_array[i]):
                next_high.extend(["higher"])
            else:
                next_high.extend(["lower"])
            #-----
            #last low
            if(low_array[i-1] >= low_array[i]):
                last_low.extend(["higher"])
            else:
                last_low.extend(["lower"])
            #-----
            #next low
            if(low_array[i+1] >= low_array[i]):
                next_low.extend(["higher"])
            else:
                next_low.extend(["lower"])
            #-----
            #last close
            if(close_array[i-1] >= close_array[i]):
                last_close.extend(["higher"])
            else:
                last_close.extend(["lower"])
            #-----
            #next close
            if(close_array[i+1] >= close_array[i]):
                next_close.extend(["higher"])
            else:
                next_close.extend(["lower"])
            #-----
            #last volume
            if(volume_array[i-1] >= volume_array[i]):
                last_volume.extend(["higher"])
            else:
                last_volume.extend(["lower"])
            #-----
            #next volume
            if(volume_array[i+1] >= volume_array[i]):
                next_volume.extend(["higher"])
            else:
                next_volume.extend(["lower"])
            
    #append the array with N/A for the first value in it
    if(i == 0):
        last_open.extend(["N/A"])
        next_open.extend(["N/A"])
        last_high.extend(["N/A"])
        next_high.extend(["N/A"])
        last_low.extend(["N/A"])
        next_low.extend(["N/A"])
        last_close.extend(["N/A"])
        next_close.extend(["N/A"])
        last_volume.extend(["N/A"])
        next_volume.extend(["N/A"])
#append the array with N/A for the last value in it
last_open.extend(["N/A"])
next_open.extend(["N/A"])
last_high.extend(["N/A"])
next_high.extend(["N/A"])
last_low.extend(["N/A"])
next_low.extend(["N/A"])
last_close.extend(["N/A"])
next_close.extend(["N/A"])
last_volume.extend(["N/A"])
next_volume.extend(["N/A"])

#run number 3
#loop through the arrays and fill in the open cells in the excel file
#new array to hold the values in each row
csv_row = []
row_number = 0
with open(csv_file, 'r') as file:
   reader = csv.reader(file)
   for row in reader:      
        #add the row to the rows array
        csv_row.extend([row])
        row_number += 1 #update the row number counter
#day.extend([row_number]) #add the last day number to the array
file.close()

#run number 4
#loop through all the csv_row values and add the needed values
#add the first row outside the loop
csv_row[0].extend(["day","last_open","next_open","last_high","next_high","last_low","next_low","last_close","next_close","last_volume", "next_volume"])
#start the for loop
for j in range(len(day)):
    if(j != 0):
        csv_row[j].extend([
                        j-1,
                        last_open[j-1],
                        next_open[j-1],
                        last_high[j-1],
                        next_high[j-1],
                        last_low[j-1],
                        next_low[j-1],
                        last_close[j-1],
                        next_close[j-1],
                        last_volume[j-1],
                        next_volume[j-1],
                        ])

#remove the second and last values in the csv_row array
#removed because of N/A values
csv_row.pop(1)
csv_row.pop(len(csv_row)-1)

#run number 5
#create a new csv file and fill it with the values from the csv_row array
with open('done_'+csv_file, 'w', newline='') as myfile:
    wr = csv.writer(myfile, quoting=csv.QUOTE_ALL)
    wr.writerows(csv_row)

print("FILE CREATED AS: "+ 'done_'+csv_file+" - YOU CAN DELETE THE OLD FILE NOW")
