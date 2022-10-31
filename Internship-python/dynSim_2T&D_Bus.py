import pssepath
pssepath.add_pssepath()
import psspy
import redirect
from openDSSAPI1 import Handler

import numpy as np
import dyntools
import os 
import matplotlib.pyplot as plt

print "\n------Initialize PSSE, run PSSE and get pu voltage at TBus------\n"

filePath='C:\Users\ce.ron\Desktop\NERC_PSSE_OpenDSS\Ron_Codes\SampleData'
rawfile = filePath+'\IEEE14bus_v32.raw'

# Initialize PSSE
psspy.psseinit(50)

# Read raw data file
psspy.read(0,rawfile)

# Run power flow
psspy.fnsl()

# Get the number of buses
ierr,nbuses = psspy.abuscount(-1,1)

# Get the pu voltages at all buses
ierr,Vpu = psspy.abusreal(-1,1,'PU')

# Location of distribution feeder
loadbus1 = 3
# Voltage at the distribution feeder load bus
Vpcc1 =  psspy.busdat(loadbus1,'PU') 

loadbus2 = 5
# Voltage at the distribution feeder load bus
Vpcc2 =  psspy.busdat(loadbus2,'PU')

# create an instance as before
dss=Handler()

# Create a mapping between distribution feeder and their IDs. Typically the ID is the name of the load bus.
mapInfo={}; mapInfo['1']={}; mapInfo['2']={} # here '1' and '2' are the feeder IDs
mapInfo['1']['filepath']=r'C:\Users\ce.ron\Desktop\NERC_PSSE_OpenDSS\Sam_Codes\SampleData\case13.dss' #make sure to use r i.e. raw string or escape backslash as \\
mapInfo['2']['filepath']=r'C:\Users\ce.ron\Desktop\NERC_PSSE_OpenDSS\Sam_Codes\SampleData\case123.dss' #change this path based on where repoLocation\data\ is located on your machine

openDSSAPIPath=r'C:\Users\ce.ron\Desktop\NERC_PSSE_OpenDSS\Ron_Codes\openDSSAPI1.py' # change this path based on where repoLocation\OpenDSS-Python\openDSSAPI.py

# Load the case
dss.load(mapInfo,openDSSAPIPath,logging=False) # this will create two subprocess, setup communication through sockets Will also load the casefiles
# Incase you need to debug, pass logging=True and see dss_out_ID.txt and dss_err_ID.txt where ID is the feeder ID.

# TBus P and Q in kw and kvar
targetS={}; targetS['1']=[50*10**3,25*10**3]; targetS['2']=[7.6*10**3,1.6*10**3]
Vpcc={}; Vpcc['1']=Vpcc1; Vpcc['2']=Vpcc2

# Set Voltage at the pcc bus of the DNetwork
dss.setV(Vpu=Vpcc)

# Scale feeder such that the feeder P,Q at PCC matches closely with targetS
K0=dss.scaleFeeder(targetS=targetS,Vpcc=Vpcc)#K0 contains feeder scaling factor K for each feeder

# Store the scaling factors in a list K
K=[]
for k,v in K0.iteritems():
    K.append(v['K'])

# Get the complex power at the feeders
S0=dss.getS() # S0 will contain S for each feeder

# Store the active and reactive demand of the feeders in lists P0 and Q0
P0 = []
Q0 = []

for k,v in S0.iteritems():
    P0.append(v['P']/1000)  
    Q0.append(v['Q']/1000)  

# Get the actual load at the transmission bus 
ierr, Sload1 = psspy.loddt2(loadbus1,'1 ','MVA','ACT')
ierr, Sload2 = psspy.loddt2(loadbus2,'1 ','MVA','ACT')

Ptbus1 = Sload1.real
Qtbus1 = Sload1.imag

Ptbus2 = Sload2.real
Qtbus2 = Sload2.imag

# Calculate the difference in load bus power and distribution system power injection
Pshunt1 = Ptbus1 - P0[0]
Qshunt1 = Qtbus1 - Q0[0]

Pshunt2 = Ptbus2 - P0[1]
Qshunt2 = Qtbus2 - Q0[1]

# Calculate compensating shunt
YPshunt1 = Pshunt1/(Vpcc1[1]**2)
YQshunt1 = -Qshunt1/(Vpcc1[1]**2)

YPshunt2 = Pshunt2/(Vpcc2[1]**2)
YQshunt2 = -Qshunt2/(Vpcc2[1]**2)

# Set P0 and Q0 (distribution system injection) in the load data for the TBus3 and TBus5

ierr = psspy.load_data_5(loadbus1,'1',[1,1,1,1,1,0,0],[P0[0],Q0[0],0,0,0,0,0,0])
ierr = psspy.load_data_5(loadbus2,'1',[1,1,1,1,1,0,0],[P0[1],Q0[1],0,0,0,0,0,0])

# Add the remaining as fixed compensating shunt
ierr = psspy.shunt_data(loadbus1,'1',1,[YPshunt1,YQshunt1])
ierr = psspy.shunt_data(loadbus2,'1',1,[YPshunt2,YQshunt2])

print "\n-----------------------End Initialization-------------------------\n"

print "Ptbus2:",Ptbus2
print "P2K:",P0[1]*K[1]
print "Pshunt2:",Pshunt2
print "YPshunt2:",YPshunt2


# CONL is done in three stages
# First stage prepare (flag = 1)
psspy.conl(-1,1,1)
# second stage convert (flag = 2)
psspy.conl(-1,1,2,[0,0],[0,100,0,100]) # 100% conversion to constant admitttance only
# third stage house keeping (flag = 3)
psspy.conl(-1,1,3) 

# Convert generators
psspy.cong()

# Read DYR file
dyrfile=filePath+'\ieee14.dyr'	
psspy.dyre_new([1,1,1,1],dyrfile)

# Add channels
psspy.chsb(0,1,[-1,-1,-1,1,13,0])       # VOLT channels for all buses	

# out file to save the channel outputs
outfilename=filePath+'\ieee14.out'
psspy.strt(outfile=outfilename)

endTime = 10.0
time = 0.5/60.0
fault_bus = 1


while time<=endTime:
                if(abs(time-1.0)<1E-4):
                    psspy.dist_bus_fault(1,1,0.0,[0.0,-1E4])

                if(abs(time-1.2)<1E-4):
                    psspy.dist_clear_fault(1)

                psspy.run(tpause=time,nprt=1000)

                Vpcc_t3 =  psspy.busdat(loadbus1,'PU')
                Vpcc_t5 =  psspy.busdat(loadbus2,'PU')

                Vpcc={}; Vpcc['1']=Vpcc_t3[1]; Vpcc['2']=Vpcc_t5[1]
				
                dss.setV(Vpu=Vpcc)

                Sdist=dss.getS()

                Pdist=[]
                Qdist=[]

                for k,v in Sdist.iteritems():
                    Pdist.append(v['P']/1000)
                    Qdist.append(v['Q']/1000)

                YadmP_t3 = Pdist[0]/(Vpcc_t3[1]*Vpcc_t3[1])
                YadmQ_t3 = -Qdist[0]/(Vpcc_t3[1]*Vpcc_t3[1])

                YadmP_t5 = Pdist[1]/(Vpcc_t5[1]*Vpcc_t5[1])
                YadmQ_t5 = -Qdist[1]/(Vpcc_t5[1]*Vpcc_t5[1])

                ierr = psspy.load_chng_5(loadbus1,'1',[1,1,1,1,1,0,0],[0,0,0,0,YadmP_t3,YadmQ_t3,0,0])
                ierr = psspy.load_chng_5(loadbus2,'1',[1,1,1,1,1,0,0],[0,0,0,0,YadmP_t5,YadmQ_t5,0,0])
                 
                time = time + (0.5/60.0)


# Load data from out file into channel file object
chnfobj = dyntools.CHNF(outfilename)

# Get the ids and data
short_title, chanid, chandata = chnfobj.get_data()

t = chandata['time']

fig = plt.figure()
fig.patch.set_facecolor('0.8')
# Plot all voltages
for i in range(1,len(chandata)):
            plt.plot(t,chandata[i])
plt.xlabel("Time")
plt.ylabel("Voltages")
axes = plt.gca()
plt.show()

# Now let us close everything gracefully (all child procs and comm)
dss.close()






