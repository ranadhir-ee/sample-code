import pssepath
pssepath.add_pssepath()
import psspy
import redirect
from openDSSAPI import OpenDSSAPI

import numpy as np
import dyntools
import os 
import matplotlib.pyplot as plt

if __name__ == '__main__':

	print "\n------Initialize PSSE, run PSSE and get pu voltage at Bus 6------\n"

	filePath='C:\Users\ce.ron\Desktop\NERC_PSSE_OpenDSS\Sam_Codes\SampleData'
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
	loadbus = 5
	# Voltage at the distribution feeder load bus
	Vpcc =  psspy.busdat(loadbus,'PU') 
		
	# Create OpenDSS instance	
	dss=OpenDSSAPI()
	baseMVA=100*10**3 # 100 MVA base. P.S. the unit is in kw hence 10**3 instead of 10**6

	# Load data file
	dss.load(r'C:\Users\ce.ron\Desktop\NERC_PSSE_OpenDSS\Sam_Codes\SampleData\IEEE123Master.dss')
	# Set PCC voltage
	dss.setV(Vpcc[1])

	# Solve power flow and get active and reactive power injection at the feeder head node
	P0,Q0,flg=dss.getS()  # P0 and Q0 are in KW.
	P0 = P0/1000.0 # Convert to MW
	Q0 = Q0/1000.0 # Convert to MVAr

        # Get the actual load at the transmission bus 
	ierr,Sload = psspy.loddt2(loadbus,'1 ','MVA','ACT')
	
	Ptbus = Sload.real
	Qtbus = Sload.imag
	
	K=dss.scaleFeeder([Ptbus*10**3, Qtbus*10**3],Vpcc=Vpcc,tol=10**-8,logging=False)# 
	print K

        # Calculate the difference in load bus power and distribution system power injection
	Pshunt = Ptbus - P0
	Qshunt = Qtbus - Q0
	
		
	# The remaining power is incorporated as compensating shunt
	# The compensating shunt power
	# Pshunt + j*Qshunt = Vpcc^2*(YPshunt - YQshunt)
	# which gives the below two equations for shunt.
	# Note the negative sign for YQshunt is because we are
	# considering admittances
	YPshunt = Pshunt/(Vpcc[1]*Vpcc[1])
	YQshunt = -Qshunt/(Vpcc[1]*Vpcc[1])

	# Set P0 and Q0 (distribution system injection) in the load data for the bus
	ierr = psspy.load_data_4(loadbus,'1 ',[1,1,1,1,1,1],[P0,Q0,0,0,0,0])

	# Add the remaining as fixed compensating shunt
	ierr = psspy.shunt_data(loadbus,'1 ',1,[YPshunt,YQshunt])

        # This power flow run is just to check whether the modifications in load
        # give the same voltages (which it should!)
	#psspy.fnsl()
        #ierr,Vpu1 = psspy.abusreal(-1,1,'PU')
        #print("  BEFORE     AFTER    DIFFERENCE")
        #for i in range(0,nbuses):   
        #        print("%9.5f %9.5f %9.5f" %(Vpu[0][i],Vpu1[0][i],Vpu[0][i]-Vpu1[0][i]))
	#print("Difference in voltage is %f",Vpcc[1]-Vpcc[1])

	
		
	print "\n------End Initialization and Begin Dynamic simulation------\n"
	
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

	# Only voltage magnitude channels ON
	psspy.chsb(0,1,[-1,-1,-1,1,13,0])       # VOLT channels for all buses	#psspy.chsb(1,0,[-1,-1,-1,1,12,0])       # FREQ channels for 138 kV buses

        # out file to save the channel outputs
	outfilename=filePath+'\ieee14.out'
	psspy.strt(outfile=outfilename)
	
	endTime = 10.0
	time = 0.5/60.0
	fault_bus = 1
	
	# Set up dynamic simulation
	
	
	P_MW=[]
	Q_MVar=[]
	t=[]
	
	dss.setDynMode(0.0083)
	
	while time <= endTime:
	

                # Apply a fault at bus 1
                if(abs(time-1.0) < 1E-4):
                        psspy.dist_bus_fault(1,1,0.0,[0.0,-1E4])

                # Clear fault after 12 cycles
                if(abs(time-1.2) < 1E-4):
                        psspy.dist_clear_fault(1)
                
		# Take a step, do not print output to stdout
		psspy.run(tpause=time,nprt=10000)

                # Get load bus voltage
                Vpcc_t =  psspy.busdat(loadbus,'PU')

		# Set the distribution feeder head node voltage
		dss.setV(Vpcc_t[1])
		
		# Run dynamic simulation at OpenDSS
		dss.runDynMode(1)

		# Run dist power flow and get sequence power
		Pdist,Qdist=dss.getSdyn() # get S
		
		
		# Scale it for units in MW/MVAr
                Pdist = Pdist/1000
                Qdist = Qdist/1000
				
				# Calculate updated admittance
                YadmP_t = Pdist/(Vpcc_t[1]*Vpcc_t[1])
                YadmQ_t = -Qdist/(Vpcc_t[1]*Vpcc_t[1])
                
                # Set admittance (distribution system injection) in the load data for the bus
                # Note: For steady-state initialization, we set constant MVA load, P0 and Q0,
                # in the load data. However, here we are setting the admittance directly because
                # the loads have been converted to admittances through psspy.conl() call.
                ierr = psspy.load_data_4(loadbus,'1 ',[1,1,1,1,1,1],[0,0,0,0,YadmP_t,YadmQ_t])
		
		P_MW.append(Pdist*K)
		Q_MVar.append(Qdist*K)
		#ierr = psspy.load_chng_5(Bus_i, '1', [1,2,2,2,1,0,0], [0,0,P*K/1000,Q*K/1000,0,0,0,0])
		t.append(time)
		time = time + (0.5/60.0)
		
	
	
	PMW = np.array(P_MW)
	QMVAR = np.array(Q_MVar)
	T = np.array(t)
	
	plt.plot(T, PMW)
	plt.xlabel("Time")
	plt.ylabel("P(MW)")
	plt.show()
	
	plt.plot(T, QMVAR)
	plt.xlabel("Time")
	plt.ylabel("Q(MVAR)")
	plt.show()
	
	
	
	'''	
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
	
	'''
	
	