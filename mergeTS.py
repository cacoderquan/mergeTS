#read from excel file
def read_excel(filename):
    wb = load_workbook(filename)
    ws = wb.active
#    print(type(ws['A2'].value))
#    print(type(ws['A2'].value < ws['A3'].value))
#    print(ws.max_row)
#    print(ws.cell(row=4, column=2).value)
    times= []
    means = []
 #   sheet_ranges = wb[name]
 #   df = pd.DataFrame(sheet_ranges.values)        
    for i in range(0, ws.max_row - 1):
        times.append(ws.cell(row = i + 2, column = 1).value)
        high = ws.cell(row = i + 2, column = 2).value
        low = ws.cell(row = i + 2, column = 3).value
        close = ws.cell(row = i + 2, column = 4).value
        means.append((high + low + close) / 3)
    return times, means 

# given two list of key values pairs and combine them into a dicionary 
def transformData(keys1, values1, keys2, values2):
    dictionary1 = dict(zip(keys1, values1))
    dictionary2 = dict(zip(keys2, values2))
    return dictionary1, dictionary2

# given two dictionaries, combine them into a single dictionary in a way that
# only the key value pairs whose keys appear in both the dictionaries will be kept. 
def combine(dictionary1, dictionary2):
    i = 0
    j = 0
    dictionary = dict()
    values = []
    keys1 = list(dictionary1.keys())
    keys2 = list(dictionary2.keys())
    values1 = list(dictionary1.values())
    values2 = list(dictionary2.values())    
    while i < len(dictionary1) and j < len(dictionary2):
        if keys1[i] == keys2[j]:
            value = [values1[i], values2[j]]
            values.append(value)
            dictionary[keys1[i]] = value
            i += 1
            j += 1
        elif keys1[i] > keys2[j]:
            j += 1
        else:
            i += 1
    return values, dictionary    


# read data 
T_times, T_means = read_excel('T1.xlsx')
TF_times, TF_means = read_excel('T2.xlsx')
# transform to dictionary
T_dict, TF_dict = transformData(T_times, T_means, TF_times, TF_means)
# combine data
values, dictionary = combine(T_dict, TF_dict)

# Y[:,0] is the price of the ten year bond future
# Y[:.1] is the price of the five year bond future
Y = np.array(values)

Rti = []    # return of ten year stock for period i 
Rtfi = []   # return of five year stock for period i 
sigmaT = []  # standard derivation of ten year stock for period 20 
sigmaTF = []  # standard derivation of five year stock for period 20 

for i in range(14, Y[:,0].size):
    # ten year return and standard derivation; calculate from the 20th element of the price array
    Treturn = (Y[i,0] - Y[i-1,0]) / Y[i-1,0]
    Rti.append(Treturn)
    a = np.asarray(Y[i-19:i,0])
    sigmaT.append(np.std(a))
    # five year return and standard derivation
    TFreturn = (Y[i,1] - Y[i-1,1]) / Y[i-1,1]
    Rtfi.append(TFreturn)
    a = np.asarray(Y[i-19:i,1])
    sigmaTF.append(np.std(a))

# create signal
signals = []
for i in range(len(Rti)):
#    signals.append(Rti[i])
    signals.append((Rti[i] / sigmaT[i]) - (Rtfi[i] / sigmaTF[i]))

plt.plot(signals)
