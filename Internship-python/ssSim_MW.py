import pssepath
pssepath.add_pssepath()
import psspy
import redirect

import numpy as np
import matplotlib.pyplot as plt
import dyntools_demo
import os 

if __name__ == '__main__':

	print "\n------Initialize PSSE, run PSSE and get pu voltage at Bus 6------\n"
	
	filePath='C:\Users\ce.ron\Desktop\NERC_PSSE_OpenDSS\Sam_Codes\SampleData'
	rawfile = filePath+'\IEEE14bus_v32.raw'
	
	# Initialize PSSE
	psspy.psseinit(50)
	# Read inputfile
	psspy.read(0,rawfile)
	
	# Run a PSSE power flow
	psspy.fnsl()
	
	Bus_i = 5 #Transmission load bus
	
	#Read loadbus voltage 
	Vpcc =  psspy.busdat(Bus_i,'PU') 
				
	print "\n------Initialize OpenDSS and scale Distribution Feeder------\n"
		
	dss=OpenDSSAPI()
	baseMVA=100*10**3 #100 MVA base. P.S. the unit is in kw hence 10**3 instead of 10**6
	
	# Load .dss files
	dss.load(r'C:\Users\ce.ron\Desktop\NERC_PSSE_OpenDSS\Ron_Codes\Steady State\PV files\IEEE123Master.dss')
	
	# Get initial P and Q requirement
	P0,Q0,flg=dss.getS()
	
	# Scale Feeder 
	K=dss.scaleFeeder([7.6*10**3, 1.6*10**3],Vpcc=Vpcc[1],tol=10**-8,logging=False)# 15 mw and 7.5 mvar (at bus 6 of ieee 14 bus system)is the target at given Vpcc value. Will display scaling progress. We will use a very small tolerance value to make the S value match exactly with the T side values. This is important because this would allow us to avoid modeling shunt to compensate for mismatch at boundary bus. P.S. the unit is in kw hence 10**3 instead of 10**6
		
	print "\n------End Initialization and Begin Co-simulation------\n"
	
	P_MW = []
	Q_MVar = []
	V_Pcc =[] 
	V_79=[]
	
	for dispatch in range(0, 24):
	    # Loadshape for 24 hours
		alpha = [0.87, 0.88, 0.86, 0.85, 0.85, 0.85, 0.86, 0.90, 0.92, 0.95, 0.95, 0.99, 0.99, 1.02, 1.05, 0.99, 0.97, 0.95, 0.98, 0.96, 0.93, 0.92, 0.91, 0.88]
		
		# Scale load according to the loadshape for each hour
		dss.scaleLoad(alpha[dispatch])
		
		# Set voltage at VSource element at OpenDSS 
		dss.setV(Vpcc[1])
		
		# Run a snapshot solve and get P and Q from D-side 		
		P,Q,flg=dss.getS() # get S
		print "\n\nOriginal feeder P={},Q={},scaled feeder P={},Q={} in PU @ Vpcc={}".format(P0,Q0,P*K/baseMVA,Q*K/baseMVA,Vpcc)
		
		# Add hourly MW requirement to the P_MW list
		P_MW.append(P*K/1000)
		Q_MVar.append(Q*K/1000)
		
		# Change load data at the load bus 
		ierr = psspy.load_data_5(Bus_i,'1',[1,1,1,1,1,0,0],[P*K/1000,Q*K/1000,0,0,0,0,0,0]) #Bus_i=5
		
        # Run a powerflow on T-side 
		psspy.fnsl()
		# Get loadbus voltage on TBus
		Vpcc =  psspy.busdat(Bus_i,'PU')
		
		V_Pcc.append(Vpcc[1])
		V_79pu=dss.getPUVol(79)
		V_79.append(V_79pu)

				
	print P_MW
		
	# Visualization
	t=[]
	for time in range(1,25):
		t.append(time)
		
	V79 = np.array(V_79)
	T = np.array(t)
	PMW = np.array(P_MW)
	
	# MW dispatch when PV is not connected to bus 79 on D-Side
	P_woPV =[6.5982, 6.6728, 6.5217, 6.4471, 6.4471, 6.4471, 6.5224, 6.8245, 6.9792, 7.2057, 7.2347, 7.5382, 7.5366, 7.7645, 7.9956, 7.5390, 7.3887, 7.2377, 7.4661, 7.3132, 7.0861, 7.0077, 6.9325, 6.7043]
	
	# Save in a .csv file 
	np.savetxt(r"C:\Users\ce.ron\Desktop\NERC_PSSE_OpenDSS\Ron_Codes\Steady State/PV_power_node79A.csv", PMW, delimiter=',')

	# Plot
	p1 = plt.bar(T-0.2, width=0.5, height=P_woPV, color='r')
	p2 = plt.bar(T+0.0, width=0.5, height=PMW, color='c')
	plt.legend((p1[0], p2[0]), ('MW dispatch without PV', 'MW dispatch with PV'))
	plt.ylabel('MW dispatch setpoints')
	plt.xlabel('hour')
	plt.title('MW dispatch setpoints at T&D interface')
	plt.show()
	