import pandas as pd

import json 



myList = []

for i in range(1,7):
    
    sname = 'S'+str(i)
    data = pd.read_excel("Timetable Workbook - SUTT Task 1.xlsx", header = 2, sheet_name=sname)


    sec = []
    for i in data['Unnamed: 6']:
        if str(i)!='nan':
            sec.append(str(i))
        else:
            try:
                sec.append(sec[-1])
            except:pass

    room = []
    for i in data['Unnamed: 8']:
        if str(i)!='nan':
            room.append(int(i))
        else:
            try:
                room.append(room[-1])
            except:
                pass

    

    instructors=[]
    to_append=[]
    for i in range(0,len(data['Unnamed: 7'])):
        
        
        if i!=1 and i!=len(data['Unnamed: 7']):
            
            if sec[i-1]==sec[i-2]:
                to_append.append(data['Unnamed: 7'][i])

            else:

                instructors.append(to_append)
                to_append=[]
                to_append.append(data['Unnamed: 7'][i])

        elif i==1:
            to_append.append(data['Unnamed: 7'][1])

    new_instructors = []
    for i in instructors:
        new_instructors.append(i)
        for j in range(len(i)-1):
            new_instructors.append(None)
    
        
    lectureUnits=data.at[0,'L']
    practicalUnits=data.at[0,'P']
    unitsUnits = data.at[0,'U']

    if lectureUnits=='-':lectureUnits=0
    elif practicalUnits=='-':practicalUnits=0
    elif unitsUnits=='-':unitsUnits=0
    

    myList.append(
        {
            
            'course_code':str(data.at[0,'Unnamed: 1']),
            'course_title':str(data.at[0,'Unnamed: 2']),
            'credits':{
                'lecture':int(lectureUnits),
                'practical':int(practicalUnits),
                'units':int(unitsUnits)

            }
            
        }
    )
    sec_type=''
    sections=[]
    grouped_keys={}

    for key, value in zip(sec,[i for i in data['Unnamed: 7']]):
        if key not in grouped_keys:
            grouped_keys[key]=[]
        grouped_keys[key].append(value)

    grouped_timetable={}
    timetable = []
    for i in data['Unnamed: 9']:
        if str(i) != 'nan':
            timetable.append(i)
        elif str(i) == 'nan':
            timetable.append(timetable[-1])
    
    midsem = []
    for i in data['Unnamed: 10']:
        if str(i) != 'nan':
            midsem.append(i)
        elif str(i) == 'nan':
            try:
                midsem.append(midsem[-1])
            except:
                for i in range(len(data['Unnamed: 10'])):
                    midsem.append('NA')
                break
    compre = []
    for i in data['Unnamed: 11']:
        if str(i) != 'nan':
            compre.append(i)
        elif str(i) == 'nan':
            try:
                compre.append(compre[-1])
            except:
                for i in range(len(data['Unnamed: 11'])):
                    compre.append('NA')
                break
    
    day=[]
    hour = []
    
    for i in timetable:
        
        if '  ' in i:
            a = i.split('  ')
        if ' ' in a[0]:
            oneDay=[]
            
            for j in a[0].split(' '):
                
                oneDay.append(j)
            
            day.append(oneDay)
        else:
            
            day.append(a[0])
        if ' ' in a[1]:
            
            oneHour=[]
            for j in a[1].split(' '):
                oneHour.append(int(j))
                oneHour.append(int(j)+2)
            minimum = min(oneHour)
            hour.append([minimum,minimum+2])
            
            
        elif ' ' not in a[1]:
            hour.append([int(a[1]),int(a[1])+1])
    
    
    newList = []
    
    
    for i in range(len(day)):
        oneList=[]
        if len(day[i])==1:
            newList.append([day[i],hour[i]])
        elif len(day[i])==2:
            for j in range(2):
                newList.append([day[i][j],hour[i]])
            # newList.append(oneList)
        elif len(day[i])==3:
            for j in range(3):
                newList.append([day[i],hour[i]])
            
            
    
    FinalFormattedTimeSlots=[]
    for d in range(len(day)):
        l=[]
        if type(day[d])==type([2]):
            for i in range(len(day[d])):
                slots=hour[d]
                output={
                    'day':day[d][i],
                    'slot':slots
                }
                l.append(output)
        else:
            slots=hour[d]
            output={
                'day':day[d],
                'slot':slots
            }
            l.append(output)

        FinalFormattedTimeSlots.append(l)
    
    done=[]
    for i in range(0,len(sec)):
        if sec[i][0]=='P':sec_type="practical"
        elif sec[i][0]=='L':sec_type="lecture"
        else:sec_type="tutorial"

        if sec[i] not in done:
            
            sections.append({             

                            "section_type":sec_type,
                            "section_number":sec[i],
                            "instructors":grouped_keys[sec[i]],
                            "room":room[i],
                            "timing":FinalFormattedTimeSlots[i],
                            "midsem":midsem[i],
                            "compre":compre[i]
                            
                            })
        done.append(sec[i])
        
    myList.append(
        
            {'sections':sections}

        
    )
    
myDict=dict()
with open('mydata.json','w') as f:
    json.dump(myList,f,indent=4,separators=(',',':'))

    


    
