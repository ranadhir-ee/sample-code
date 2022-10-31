from __future__ import division
from subprocess import Popen
import win32com.client,os,sys,linecache
import numpy as np,socket,copy,shlex,json

#========================================================
#===================WORKER CLASS========================
#========================================================
class Worker():
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
			self.__enumerations()# populate enumerations
		else:
			print "DSS Failed to Start"
			self.flg_startComm=-1 # fail
		
		return None

	#===================LOAD A CASE========================
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
			PrintException()
			
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
			PrintException()
			
	#===================GET LOADS========================
	def getLoads(self):
		"""Get the load setting for every load in the system. Please note that this 
		is not the actual load consumed. To get that you need a meter at the load bus.
		All values are reported in KW and KVar"""
		try:
			S={};S['P']={};S['Q']={}
			self.Loads.First # start with the first load in the list
			for n in range(0,self.Loads.Count):
				S['P'][self.Loads.Name],S['Q'][self.Loads.Name]=\
				self.Loads.kw,self.Loads.kvar
				self.Loads.Next # move to the next load in the system

			return S
		except:
			PrintException()
		
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
				self.Loads.kw=self.S0['P'][self.Loads.Name]*scale
				self.Loads.kvar=self.S0['Q'][self.Loads.Name]*scale
				self.Loads.Next # move to the next load in the system
		except:
			PrintException()

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
			PrintException()

#===============================SET VOLTAGE===============================
	def setV(self,Vpu,Vang=0,pccName='Vsource.source'):
		"""Sets the PCC voltage."""
		try:
			self.changeObj([[pccName,'pu',Vpu,'set'],[pccName,'angle',Vang,'set']])
		except:
			PrintException()

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
				self.Circuit.ActiveCktElement.SeqPowers[3]*-1####TODO: is this the best way?
			else:
				P,Q=None,None
			return P,Q,self.Solution.Converged
		except:
			PrintException()

#===============================SCALE FEEDER===============================
	def scaleFeeder(self,targetS,Vpcc,tol=10**-6,zipv=None,logging=False):
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
			PrintException()

#===================GET PCC S========================
	def __enumerations(self):
		try:
			self.enum={}; mode=self.enum['solutionMode']={}
			
			modes=['Snap','Daily','Yearly','MI','LD1','Peakday','DUtycycle','DIrect','MF',\
			'FaultStudy','M2','M3','LD2','Autoadd','Dynamic','Harmonic']
			
			for n in range(0,len(modes)):
				mode[modes[n]]=n
		except:
			PrintException()

#========================================================
#===================HANDLER CLASS========================
#========================================================
class Handler():
	"""A handler class that allows managing multiple distribution feeders.

	Typical usage:

	foo=Handler()
	mapInfo={};mapInfo['1']={} mapInfo['1']['filePath']=path/to/case13.dss
	mapInfo['2']={}; mapInfo['2']['filepath']=path/to/case123.dss
	foo.load(mapInfo)# will instantiate two workers to run 13 and 123 node feeders.
	Vpu={}; Vpu['1']=1.0; Vpu['2']=0.98
	reply=foo.setV(Vpu)
	reply=foo.getS()
	
	reply will be of format,
	reply['1'] and reply['2']
	where reply['1'] (and '2') will have P,Q,convergenceFlg as key,value pair
	foo.close()# do not forget to close the handler. The close will signal all the
	child procs to quit gracefully.
	"""
	####TODO: The current approach will result in too many context switches. 
	# Initial experiments resulted in about 60% percent avg proc use across all cores.
	# Could we limit the number of procs
	# and combine serveral worker calls to one process? 
	# COM binding is currently the limiting factor.
	#
	####TODO: Add support for changeObj, solve, scaleLoad, getVoltage methods of Worker
	# Add memory management using psutil. Better standard cross-platform lib?
	# Add methods to check if child procs quit gracefully, if not use subprocess to
	# terminate the child procs.

	def __init__(self):
		self.__startServer()
		self.BUFFER_SIZE=1024*8
		return None

#===================START SERVER========================
	def __startServer(self,host='127.0.0.1',portNum=11000):
		"""Initialize the server for communication with worker."""
		try:
			self.s = socket.socket(socket.AF_INET,socket.SOCK_STREAM)
			self.s.bind((host,portNum))
			self.s.listen(0)
		except:
			PrintException()

	#===================INITIALIZE AND LOAD CASE========================
	def load(self,mapInfo,openDSSAPIPath,logging=False):
		"""mapInfo should be of the following format,
		{feederID:{"filePath":fpath},...}
		Where feederID is the user defined ID for distribution feeder.
		If there are n feederIDs then n process is launched.
		
		openDSSAPIPath - path/to/openDSSAPI.py
		"""
		try:
			numConn=len(mapInfo.keys())
			self.mapInfo={}
			self.mapInfo=copy.deepcopy(mapInfo)

			for entry in self.mapInfo.keys():
				if logging:
					self.mapInfo[entry]['f_out']=open('dss_out_'+entry+'.txt','w')
					self.mapInfo[entry]['f_err']=open('dss_err_'+entry+'.txt','w')
				else:
					self.mapInfo[entry]['f_out']=self.mapInfo[entry]['f_err']=\
					open(os.devnull,'w')

				# start a subprocess asynchronously, one at a time
				self.mapInfo[entry]['proc']=Popen(shlex.split("python "+'"'+openDSSAPIPath+'"'),\
				stdout=self.mapInfo[entry]['f_out'],\
				stderr=self.mapInfo[entry]['f_err'])

				#accept connection from worker
				####TODO: Need to make handshake with worker to make sure the right
				# process is connecting.
				self.mapInfo[entry]['conn']=self.s.accept()

				# now load the case
				msg={}; msg['filepath']=self.mapInfo[entry]['filepath']
				msg['method']='load'
				self.mapInfo[entry]['conn'][0].send(json.dumps(msg))# send msg
				reply=json.loads(self.mapInfo[entry]['conn'][0].recv(self.BUFFER_SIZE))
		except:
			PrintException()

	#===================CLOSE========================
	def close(self):
		try:
			# close workers
			for entry in self.mapInfo.keys():
				self.mapInfo[entry]['conn'][0].send('{"COMM_END":1}')
				shutdownReply=self.mapInfo[entry]['conn'][0].recv(self.BUFFER_SIZE)
				
			#close server
			self.s.close() # close server
			
			# kill child processes, if necessary
			childPID=[]
			for entry in self.mapInfo.keys():
				if self.mapInfo[entry]['proc'].poll() is None:# process is running
					Popen.kill(self.mapInfo[entry]['proc'])# At this point we no longer need child process
					childPID.append(self.mapInfo[entry]['proc'].pid)

			# Popen.kill will not work always. Implemented psutil based logic.
			# Get the clients that connect to handler server, check for PID, if the PID intersects with
			# subprocess PID then we know that this is a child process. Then force kill the child process
			# using taskkill through os.system. Downside: Needs a non-standard python lib psutil.
			try:
				import psutil
				netConn=psutil.net_connections()
				for entry in netConn:
					if len(entry.raddr)>0:
						if entry.raddr[1]==11000:
							if entry.pid is not None and entry.pid in childPID and \
							entry.status.lower()=='established':
								os.system('taskkill /pid {} /f >> opendss_pkill_log.txt 2>&1'.format(entry.pid))
			except ImportError:# library not installed
				pass
		except:
			PrintException()

	#===================SCALE FEEDER========================
	def scaleFeeder(self,targetS,Vpcc,tol=10**-6):
		"""targetS, Vpcc and tol should be dictionary of the following format,
		{feederID:value}"""
		try:
			# assume default values if not provided
			if not isinstance(Vpcc,dict):
				Vpcc_default=Vpcc; Vpcc={}
				for entry in targetS.keys():
					Vpcc[entry]=Vpcc_default
			if not isinstance(tol,dict):
				tol_default=tol; tol={}
				for entry in targetS.keys():
					tol[entry]=tol_default

			for entry in targetS.keys():# first send to all to allow computation to be run concurrently
				msg={}; msg['method']='scaleFeeder'; msg['targetS']=targetS[entry]
				msg['Vpcc']=Vpcc[entry]; msg['tol']=tol[entry]
				self.mapInfo[entry]['conn'][0].send(json.dumps(msg))# send msg

			replyMsg={}
			for entry in targetS.keys():# now receive replies
				replyMsg[entry]=json.loads(self.mapInfo[entry]['conn'][0].recv(self.BUFFER_SIZE))

			return replyMsg
		except:
			PrintException()

	#===================SET VOLTAGE========================
	def setV(self,Vpu,Vang=0,pccName='Vsource.source'):
		try:
			# assume default values if not provided
			if not isinstance(Vang,dict):
				Vang_default=Vang; Vang={}
				for entry in Vpu.keys():
					Vang[entry]=Vang_default
			if not isinstance(pccName,dict):
				pccName_default=pccName; pccName={}
				for entry in Vpu.keys():
					pccName[entry]=pccName_default

			# send and recv
			for entry in Vpu.keys():# first send to all to allow computation to be run concurrently
				msg={}; msg['method']='setV'; msg['Vpu']=Vpu[entry]
				msg['Vang']=Vang[entry]; msg['pccName']=pccName[entry]
				self.mapInfo[entry]['conn'][0].send(json.dumps(msg))# send msg

			replyMsg={}
			for entry in Vpu.keys():# now receive replies
				replyMsg[entry]=json.loads(self.mapInfo[entry]['conn'][0].recv(self.BUFFER_SIZE))
		except:
			PrintException()

	#===================GET PCC S========================
	def getS(self,pccName='Vsource.source'):
		try:
			# assume default values if not provided
			if not isinstance(pccName,dict):
				pccName_default=pccName; pccName={}
				for entry in self.mapInfo.keys():
					pccName[entry]=pccName_default

			# send and recv
			for entry in self.mapInfo.keys():# first send to all to allow computation to be run concurrently
				msg={}; msg['method']='getS'; msg['pccName']=pccName[entry]
				self.mapInfo[entry]['conn'][0].send(json.dumps(msg))# send msg

			replyMsg={}
			for entry in self.mapInfo.keys():# now receive replies
				replyMsg[entry]=json.loads(self.mapInfo[entry]['conn'][0].recv(self.BUFFER_SIZE))

			return replyMsg
		except:
			PrintException()
	
	#===================GET scaleLoad========================
	def scaleLoad(self,scale):
		try:
			#assume default values if not provided
			if not isinstance(scale,dict):
				scale_default=scale; scale={}
				for entry in scale.keys():
					scale[entry]=scale_default

			#send and recv
			for entry in self.mapInfo.keys():# first send to all to allow computation to be run concurrently
				msg={}; msg['method']='scaleLoad'; msg['scale']=scale[entry]; 
				self.mapInfo[entry]['conn'][0].send(json.dumps(msg))# send msg

			replyMsg={}
			for entry in self.mapInfo.keys():# now receive replies
				replyMsg[entry]=json.loads(self.mapInfo[entry]['conn'][0].recv(self.BUFFER_SIZE))
			
			return replyMsg	
		except:
			PrintException()
	

	#===================GET getLoads========================
	def getLoads(self):
		try:
		
			# send and recv
			for entry in self.mapInfo.keys():# first send to all to allow computation to be run concurrently
				msg={}; msg['method']='getLoads'; 
				self.mapInfo[entry]['conn'][0].send(json.dumps(msg))# send msg

			replyMsg={}
			for entry in self.mapInfo.keys():# now receive replies
				replyMsg[entry]=json.loads(self.mapInfo[entry]['conn'][0].recv(self.BUFFER_SIZE))

			return replyMsg
		except:
			PrintException()
	

#===================PRINT EXCEPTION========================
def PrintException():
	 exc_type, exc_obj, tb = sys.exc_info()
	 f = tb.tb_frame
	 lineno = tb.tb_lineno
	 filename = f.f_code.co_filename
	 linecache.checkcache(filename)
	 line = linecache.getline(filename, lineno, f.f_globals)
	 print "Exception in Line {}".format(lineno)
	 print "Error in Code: {}".format(line.strip())
	 print "Error Reason: {}".format(exc_obj)

#===================RUN AS SCRIPT========================

if __name__=="__main__":
	try:
		dss=Worker()
	
		BUFFER_SIZE=1024*8
		# connect to handler
		c = socket.socket(socket.AF_INET,socket.SOCK_STREAM)
		c.connect(('127.0.0.1',11000))

		comm_end=0
		while comm_end==0:
			msg=json.loads(c.recv(BUFFER_SIZE))# expect the msg to be of json format
			if not msg.has_key('COMM_END'):
				# process request
				####TODO: perhaps change everything in worker api to kwargs to allow 
				# for a more generic call?
				replyMsg={}
				if msg['method'].lower()=='load':
					dss.load(msg['filepath'])
				elif msg['method'].lower()=='scalefeeder':
					replyMsg['K']=dss.scaleFeeder(targetS=msg['targetS'],Vpcc=msg['Vpcc'],\
					tol=msg['tol'])
				elif msg['method'].lower()=='setv':
					dss.setV(Vpu=msg['Vpu'],Vang=msg['Vang'],pccName=msg['pccName'])
				elif msg['method'].lower()=='gets':
					replyMsg['P'],replyMsg['Q'],replyMsg['convergenceFlg']=\
					dss.getS(pccName=msg['pccName'])
				elif msg['method'].lower()=='scaleLoad':
					replyMsg=dss.scaleLoad()
				elif msg['method'].lower()=='getloads':
					replyMsg['S']=dss.getLoads()
                    
				c.send(json.dumps(replyMsg))# reply back to handler

			elif msg.has_key('COMM_END'):
				comm_end=1
				c.send(json.dumps({"shutdown":1}))# reply back to handler
				c.shutdown(0)
				c.close() # close comm with server
	except:
		PrintException()
