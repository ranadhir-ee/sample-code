import win32com.client
import os
import sys
import linecache
import numpy as np

class OpenDSSAPI():
	"""API class for openDSS."""
	def __init__(self):
		####TODO: Early binding is supposed to speed things up wrt COM, however, it does
		# not seem to do that. In addition, Loads.kw is not found in the variable 
		# dictionary. Need to check why this happens. For now, early binding is disabled.

		self.engine = win32com.client.Dispatch("OpenDSSEngine.DSS")

		if self.engine.Start("0") == True:#DSS started OK
			self.flg_startComm=1 # pass
			self.Circuit=self.engine.ActiveCircuit

			self.Text,self.Solution,self.CktElement,self.Bus,self.Meters,\
			self.PDElement,self.Loads,self.Lines,self.Transformers,\
			self.Monitors = self.engine.Text,\
			self.Circuit.Solution,self.Circuit.ActiveCktElement,\
			self.Circuit.ActiveBus,self.Circuit.Meters,self.Circuit.PDElements,\
			self.Circuit.Loads,self.Circuit.Lines,self.Circuit.Transformers,\
			self.Circuit.Monitors
			self.startingDir=os.getcwd()
		else:
			print "DSS Failed to Start"
			self.flg_startComm=-1 # fail
		
		return None

#===================LOAD A CASE========================
	"""
	def load(self,fname):
		try:
			# Always a good idea to clear the DSS before loading a new circuit
			self.engine.ClearAll()

			self.fname=fname
			print "loaded: ",self.fname
			self.Text.Command = "compile [" + self.fname + "]"
			os.chdir(self.startingDir)# openDSS sets dir to data dir, revert back
			
			# once loaded find the base load
			self.S0=self.getLoads()
		except:
			self.PrintException()
	"""		
	def load(self,fname):
		try:
			# Always a good idea to clear the DSS before loading a new circuit
			#self.engine.ClearAll()

			self.fname=fname
			print "loaded: ",self.fname
#			self.Text.Command = "compile [" + self.fname + "]"
			self.Text.Command = "Redirect [" + self.fname + "]"
			os.chdir(self.startingDir)# openDSS sets dir to data dir, revert back
			
			# once loaded find the base load
			self.S0=self.getLoads()
		except:
			PrintException()

#===================SOLVE========================
	def solve(self,solveCmd="solve"):
		"""Solves the active circuit.
		Input: solveCmd -- user can control specific solve operation through this."""
		try:
			self.Text.Command = solveCmd
			print "Convergence status: ",self.Solution.Converged
			return None
		except:
			self.PrintException()
			
#===================GET VOLTAGE OF ALL NODES========================
	def getVoltage(self):
		"""Needs to be called after solve. Gathers Voltage (complex) of all buses into
		 a,b,c phases and populates a dictionary and returns it."""
		try:
			Voltage={}; entryMap={}; entryMap['1']='a'; entryMap['2']='b';
			entryMap['3']='c'

			for n in range(0,self.Circuit.NumBuses):
				Voltage[self.Circuit.Buses(n).Name]={}
				V=self.Circuit.Buses(n).Voltages

				count=0
				for entry in self.Circuit.Buses(n).Nodes:
					Voltage[self.Circuit.Buses(n).Name][entryMap[str(int(entry))]]=\
					V[count]+1j*V[count+1]
					count+=2

			return Voltage
		except:
			self.PrintException()
	
#===============================FOO===============================
	def setV(self,Vpu,pccName='Vsource.source'):
		"""Sets the PCC voltage."""
		try:
			self.Text.Command='Edit {} pu={}'.format(pccName,Vpu)
		except:
			self.PrintException()

#===================GET PCC S========================
	def getS(self,pccName='Vsource.source'):
		"""Gets the positive seq complex power requirement at pcc bus.
		Input: pccName - name of the pcc bus"""
		try:
			self.Solution.Solve()
			if self.Solution.Converged:
				self.Circuit.SetActiveElement(pccName)
				# -ve sign convention
				P,Q=self.Circuit.ActiveCktElement.SeqPowers[2]*-1,\
				self.Circuit.ActiveCktElement.SeqPowers[3]*-1
				
				
			else:
				P,Q=None,None
			return P,Q,self.Solution.Converged
		except:
			self.PrintException()

#===============================FOO===============================
	def scaleFeeder(self,targetS,Vpcc=1.0,tol=10**-3,zipv=None,logging=False):
		"""Will scale feeder such that the feeder P,Q at PCC matches closely with targetS.
		Additionally Will convert loads to zip loads if requested.
        
		Input: targetS -- should be in [kw,kvar] format
		Vpcc - voltage magnitude at PCC at which targetS is to be achieved.
		tol -- tolerance for power factor match
		zip -- convert loads to zip load.
		logging -- print helpful info.

		Output: K -- Scaling factor. None is returned if power flow diverged.
		"""
		####TODO:i) add zip load change ii) power flow diverges for 8500 node test case. Hence, alg fails.
		try:
			# match PF
			dQ=10 # in kvar
			maxIter=100
			self.Solution.Solve()
			convergedFlg=self.Solution.Converged
			
			self.setV(Vpcc)
			P,Q,convergedFlg=self.getS()
			if convergedFlg:
				targetPF=np.cos(np.angle(np.complex(targetS[0],targetS[1])))
				currentPF=np.cos(np.angle(np.complex(P,Q)))

			dQChange=0; iterCount=0
			while abs(currentPF-targetPF)>tol and iterCount<maxIter and convergedFlg:
				if dQChange>abs(currentPF-targetPF):# reduce dQ if needed
					dQ=dQ/2
				self.Loads.First # start with the first load in the list
				for n in range(0,self.Loads.Count):
					if currentPF>targetPF and self.Loads.kvar!=0: # increase Q
						self.Loads.kvar=self.Loads.kvar+dQ
					elif currentPF<targetPF and self.Loads.kvar!=0: # reduce Q
						self.Loads.kvar=self.Loads.kvar-dQ
					self.Loads.Next # move to the next load in the system

				P,Q,convergedFlg=self.getS()
				if convergedFlg:
					currentPFNew=np.cos(np.angle(np.complex(P,Q)))
					dQChange=abs(currentPF-currentPFNew)
					currentPF=currentPFNew
					iterCount+=1
				if logging:
					print "iteration no={}, currentPF={}, targetPF={}".format(iterCount,currentPF,targetPF)

			# find scaling
			if convergedFlg:
				K=targetS[0]/P # scaling
			else:
				print "power flow solve diverged."
				K=None

			return K
		except:
			self.PrintException()
#===================GET LOADS========================
	def getLoads(self):
		"""Get the load setting for every load in the system. Please note that this 
		is not the actual load consumed. To get that you need a meter at the load bus.
		All values are reported in KW and KVar"""
		try:
			S={};S['P']={};S['Q']={}
			iLoads=self.Loads.First # start with the first load in the list
			#for n in range(0,self.Loads.Count):
			while iLoads:
				S['P'][self.Loads.Name] = self.Loads.kW
				S['Q'][self.Loads.Name] = self.Loads.kvar
				
				iLoads=self.Loads.Next # move to the next load in the system

			return S
		except:
			self.PrintException()
		
#===================LOAD SHAPE========================
	def scaleLoad(self,scale):
		"""Sets the load shape by scaling each load in the system with a scaling
			factor scale.
			Input: scale -- A scaling factor for loads such that P+j*Q=scale*(P0+j*Q0)
			P.S. loadShape should be called at every dispatch i.e. only load(t) is set.
			"""
		try:
			self.Loads.First # start with the first load in the list
			for n in range(0,self.Loads.Count):
				self.Loads.kW=self.S0['P'][self.Loads.Name]*scale
				self.Loads.kvar=self.S0['Q'][self.Loads.Name]*scale
				self.Loads.Next # move to the next load in the system
		except:
			self.PrintException()
			
#===================SET/GET OBJECT VALUE========================
	def changeObj(self,objData):
		"""set/get an object property.
		Input: objData should be a list of lists of the format,
		[[objName,objProp,objVal,flg],...]

		objName -- name of the object.
		objProp -- name of the property.
		objVal -- val of the property. If flg is set as 'get', then objVal is not used.
		flg -- Can be 'set' or 'get'

		P.S. In the case of 'get' use a value of 'None' for objVal. The same object i.e.
		objData that was passed in as input will have the result i.e. objVal will be
		updated from 'None' to the actual value.
		
		Sample call: self.changeObj([['PVsystem.pv1','kVAr',25,'set']])
		self.changeObj([['PVsystem.pv1','kVAr','None','get']])
		"""
		try:
			for entry in objData:
				self.Circuit.setActiveElement(entry[0])# make the required element as active element
				if entry[-1]=='set':
					self.CktElement.Properties(entry[1]).Val=entry[2]
				elif entry[-1]=='get':
					entry[2]=self.CktElement.Properties(entry[1]).Val
		except:
			self.PrintException()

#===================GET PCC S========================
	def getSPCC(self,line_pccBus2next):
		"""Gets the positive seq complex power requirement at
		pcc bus.
		Input: line_pccBus2next -- name of the line connecting 
		pcc bus to the next node."""
		try:
			self.Circuit.SetActiveElement('Line.'+line_pccBus2next)
			S=self.Circuit.ActiveCktElement.SeqPowers
			P1=S[2]; Q1=S[3]####TODO: is this the best way?
			return P1,Q1 # positive seq P and Q
		except:
			self.PrintException()

#===================GET kVBase========================
	def getkVBase(self, bus_num):
		"""Gets the kvBase at any bus.
		Input: Bus number as defined in the DSS script.
		Usage: 
		Example: dss.getkVBase(150)
		
		Note: 
		Getting kVBase = 0.24 for all the buses.
		
		OpenDSS works in actual volts and amps. Voltage 
		bases are mostly used for output reports. 
		If you use the SetkVBase command you can set it 
		with either kVLL or kVLN. Then CalcVoltageBases 
		command is used to automatically guess the desired 
		voltage base from the Set VoltageBase option array. 
		
		More explanation here: 
		https://sourceforge.net/p/electricdss/discussion/beginners/thread/ff4f4f1c/?limit=25#f58f
		"""
		
		try:
			self.Circuit.SetActiveBus(str(bus_num))
			B=self.Circuit.ActiveBus.kVBase
			return B # returns kVBase 
			
		
		except:
			self.PrintException()
			
#===================GET SeqCurrents========================
	def getSeqI(self, pu):
		"""Get sequence currents at the pcc bus.
		Input: line_pccBus2next -- name of the line connecting 
		pcc bus to the next node; per-unit voltage(pu)
		"""
		try:
			self.Circuit.SetActiveElement('Line.115')     # Voltage Source at Bus 150. L.115 is connected between buses 150 and 149
			self.Text.Command = "Edit Vsource.Source" + " pu=" + str(pu)
			self.solve()
			I=self.Circuit.ActiveCktElement.SeqCurrents
			I0=I[0]; I1=I[1]; I2=I[2] 			
			return I[1] # positive sequence current
		except:
			self.PrintException()

#===================Populate PV===================================
	def placePV(self, PV_name, phase_num, bus_num, irr, kV, kVA, Pmpp):
		"""
                Input: 
                PV_name -- name of the PV system, phase_num -- phase# of the bus
                on which the PV is placed, bus_num -- bus name, irr -- irradiance    kV -- voltage level,
                kV, kVA, Pmpp.
		"""
		try:
			self.Text.Command = "New PVSystem." + str(PV_name) + " phases=" + str(phase_num) + " bus1=" + str(bus_num) + " irradiance=" + str(irr) + " kV=" + str(kv) + " kVA=" + str(kVA) + " Pmpp=" + str(Pmpp) 
		except:
			self.PrintException()

#===================PLACE Monitors===================================
	def placeMon(self, Mon_name, Line_name, Ter_num, mode):
		"""
                Input: Mon_name -- name of the monitor, 
					   Line_name -- the line on which you want to place the monitor
					   Ter_num -- terminal 1 or 2 of the line
					   Mode -- 	0: Voltages and Currents
								1: Powers
								2: Transformer taps (Transformer elements only)
								3: State Variables (PCElements only)
								4: Flicker (Pst) of Voltages only (10 min simulation required) 
						
						Monitor Objects can capture sequence quantities directly by a bitmask
						adder to the Mode number. Combine with adders below to achieve other
						results for terminal quantities:
						+16 = Sequence quantities
						+32 = Magnitude only
						+64 = Positive sequence only or avg of all phases
						
						e.g., Mode=17 will return sequence powers 
		"""
		try:
			self.Text.Command = "New Monitor." + str(Mon_name) + " Element=Line." + str(Line_name) + " Terminal=" + str(ter_num) +  " Mode=" + str(mode) + "PPolar=No" 
		except:
			self.PrintException()
			
#===================Simulate Bus Fault===================================
	def busFault(self, Fault_name, Bus_name, phase_num, time_step, step_size):
		
		"""		Add a SLG fault to phase "phase_num" of a Bus (phases=1 by default) 
				e.g: 
                New Fault.F1 Bus1=MyBus.2
				Solve Mode=Dynamic Number=1 Stepsize=0.00001 
				
				Solve 1 timestep in Dynamic mode at a tiny timestep.
				This captures Generator contribution to fault.
				
				Input: 
				Fault_name -- name of the fault, 
				Bus_name -- bus on which fault is applied, 
				phase_num -- phase of the bus on which fault is applied, 
				time_step --  
				step_size --
		"""
		try:
			self.Text.Command = "New Fault." + str(Fault_name) + " Bus1=" + str(Bus_name) + "." + str(phase_num) + "\n"
			self.Text.Command = "Solve Mode=Dynamic Number=" + str(time_step) + " stepsize=" + str(step_size)
			self.solve() 
			
		except:
			self.PrintException()

#===================GET PCC S========================
	def __enumerations(self):
		try:
			self.enum={}; mode=self.enum['solutionMode']={}
			
			modes=['Snap','Daily','Yearly','MI','LD1','Peakday','DUtycycle','DIrect','MF',\
			'FaultStudy','M2','M3','LD2','Autoadd','Dynamic','Harmonic']
			
			for n in range(0,len(modes)):
				mode[modes[n]]=n
		except:
			self.PrintException()
			
#===================GET puVoltages========================
	def getPUVol(self,bus_name):
		"""Get average pu voltage at any bus.
		Input: Input bus name of the circuit. use dss.solve() before using the function.	
		
		## this function may be modified to get individual phase voltages
		"""
		try:
			self.Circuit.SetActiveElement(str(bus_name))
			temp_vol=self.Bus.puVmagAngle
            
			if len(temp_vol) == 2:
				bus_vol = (self.Bus.puVmagAngle[0])
			if len(temp_vol) == 4:
				bus_vol = ((self.Bus.puVmagAngle[0]+self.Bus.puVmagAngle[2])/2)
			if len(temp_vol) == 6:
				bus_vol = ((self.Bus.puVmagAngle[0]+self.Bus.puVmagAngle[2]+self.Bus.puVmagAngle[4])/3)
			
			return bus_vol
		except:
			self.PrintException()
	#===================Populate PV===================================
	def placePV(self, PV_name, phase_num, bus_num, irr, kV, kVA, Pmpp):
		"""
                Input: 
                PV_name -- name of the PV system, phase_num -- phase# of the bus
                on which the PV is placed, bus_num -- bus name, irr -- irradiance    kV -- voltage level,
                kV, kVA, Pmpp.
		"""
		try:
			self.Text.Command = "New PVSystem." + str(PV_name) + " phases=" + str(phase_num) + " bus1=" + str(bus_num) + " irradiance=" + str(irr) + " kV=" + str(kv) + " kVA=" + str(kVA) + " Pmpp=" + str(Pmpp) 
		except:
			self.PrintException()


#===================PRINT EXCEPTION========================
	def PrintException(self):
		 exc_type, exc_obj, tb = sys.exc_info()
		 f = tb.tb_frame
		 lineno = tb.tb_lineno
		 filename = f.f_code.co_filename
		 linecache.checkcache(filename)
		 line = linecache.getline(filename, lineno, f.f_globals)
		 print "Exception in Line {}".format(lineno)
		 print "Error in Code: {}".format(line.strip())
		 print "Error Reason: {}".format(exc_obj) 

#+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'''	
if __name__ == '__main__':
	from openDSSAPI import OpenDSSAPI

	#dir_path = os.path.dirname(os.path.realpath(__file__))
	dir_path = r'C:\Program Files\OpenDSS\IEEETestCases\123Bus'
	dss=OpenDSSAPI()
	dss.load(dir_path + '/IEEE123Master.dss')
	
	dss.solve()
	#V=dss.getVoltage()
	#L=dss.getLoads()
	#print V
	#print L
	#puV = dss.getPUVol(60)
	#print puV
	
	
	I=dss.getSeqI(0.95)
	print I
	I=dss.getSeqI(1.00)
	print I
	I=dss.getSeqI(1.05)
	print I
	
	B=dss.getLoads()
	print B
'''


