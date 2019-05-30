#coding=utf-8
import time,random,threading,numpy
import win32com.client

pan=[-1]*480
gailv1=[0]*480
dian=[]
dm=win32com.client.Dispatch('dm.dmsoft')
dm.leftclick()

def zuobiao(x,y):
	x=int((x-1342)/18)
	y=int((y-80)/18)
	return x,y
def end():
	intX=intY=0
	ret = dm.FindPic(1332,71,1886,370,"zd.bmp","000000",0.9,0,intX,intY)
	if ret[0]!=-1:
		os.exit()
def find(pan):	
	ret=dm.findpicex(1342,80,1883,369,"0.bmp","000000",0.6,0)
	ret=ret.split("|")
	if ret!=[""]:
		for i in ret:
			i=i.split(",")
			x,y=zuobiao(int(i[1]),int(i[2]))
			#dm.moveto(int(i[1]),int(i[2]))
			#time.sleep(0.3)
			m=y*30+x
			pan[m]=0
	ret=dm.findpicex(1342,80,1883,369,"1.bmp","000000",0.6,0)
	ret=ret.split("|")
	if ret!=[""]:
		for i in ret:
			i=i.split(",")
			x,y=zuobiao(int(i[1]),int(i[2]))
			#dm.moveto(int(i[1]),int(i[2]))
			#time.sleep(0.3)
			m=y*30+x
			pan[m]=1

	ret=dm.findpicex(1342,80,1883,369,"2.bmp","000000",0.6,0)
	ret=ret.split("|")
	if ret!=[""]:
		for i in ret:
			i=i.split(",")
			x,y=zuobiao(int(i[1]),int(i[2]))
			#dm.moveto(int(i[1]),int(i[2]))
			#time.sleep(0.3)
			m=y*30+x
			pan[m]=2
	ret=dm.findpicex(1342,80,1883,369,"3.bmp","000000",0.6,0)
	ret=ret.split("|")
	if ret!=[""]:
		for i in ret:
			i=i.split(",")
			x,y=zuobiao(int(i[1]),int(i[2]))
			#dm.moveto(int(i[1]),int(i[2]))
			#time.sleep(0.3)
			m=y*30+x
			pan[m]=3
	ret=dm.findpicex(1342,80,1883,369,"4.bmp","000000",0.6,0)
	ret=ret.split("|")
	if ret!=[""]:
		for i in ret:
			i=i.split(",")
			x,y=zuobiao(int(i[1]),int(i[2]))
			#dm.moveto(int(i[1]),int(i[2]))
			#time.sleep(0.3)
			m=y*30+x
			pan[m]=4
	ret=dm.findpicex(1342,80,1883,369,"5.bmp","000000",0.6,0)
	ret=ret.split("|")
	if ret!=[""]:
		for i in ret:
			i=i.split(",")
			x,y=zuobiao(int(i[1]),int(i[2]))
			#dm.moveto(int(i[1]),int(i[2]))
			#time.sleep(0.3)
			m=y*30+x
			pan[m]=5
	'''ret=dm.findpicex(1332,71,1886,370,"6.bmp","000000",0.7,0)
	ret=ret.split("|")
	if ret!=[""]:
		for i in ret:
			i=i.split(",")
			x,y=zuobiao(int(i[1]),int(i[2]))
			m=y*18+x+1
			pan[m]=6'''
	if pan==[-1]*480:
		x=random.randint(10,20)
		y=random.randint(5,11)
		dm.moveto(1342+x*18+5,80+y*18+5)
		time.sleep(2)
		dm.leftclick()
		time.sleep(0.1)
		dm.moveto(1342,70)	
def jisuan(pan):
	lei1=[]
	lei2=[]
	lei3=[]
	lei4=[]
	global new,dian,gailv1	
	m=0
	for i in pan:
		n=[]
		zd=[]
		if i==-1 or i==0 or i==-2:
			pass
		else:
			x=m%30
			y=int(m/30)
			if 0<=x-1<=29 and 0<=y-1<=15:			
				if pan[m-31]==-1:
					n.append(m-31)
				if pan[m-31]==-2:
					zd.append(m-31)
			if 0<=x<=29 and 0<=y-1<=15:			
				if pan[m-30]==-1:
					n.append(m-30)
				if pan[m-30]==-2:
					zd.append(m-30)	
			if 0<=x+1<=29 and 0<=y-1<=15:			
				if pan[m-29]==-1:
					n.append(m-29)	
				if pan[m-29]==-2:
					zd.append(m-29)	
			if 0<=x-1<=29 and 0<=y<=15:			
				if pan[m-1]==-1:
					n.append(m-1)
				if pan[m-1]==-2:
					zd.append(m-1)	
			if 0<=x+1<=29 and 0<=y<=15:			
				if pan[m+1]==-1:
					n.append(m+1)
				if pan[m+1]==-2:
					zd.append(m+1)	
			if 0<=x-1<=29 and 0<=y+1<=15:			
				if pan[m+29]==-1:
					n.append(m+29)
				if pan[m+29]==-2:
					zd.append(m+29)	
			if 0<=x<=29 and 0<=y+1<=15:			
				if pan[m+30]==-1:
					n.append(m+30)
				if pan[m+30]==-2:
					zd.append(m+30)	
			if 0<=x+1<=29 and 0<=y+1<=15:			
				if pan[m+31]==-1:
					n.append(m+31)
				if pan[m+31]==-2:
					zd.append(m+31)
			if len(n)+len(zd)==i:
				for i in n:
					pan[i]=-2					
			elif len(zd)==i:
				for i in n:
					dian.append(i)	
					pan[i]=0
			else:
				if i-len(zd)==1:
					lei1.append(n)		
				elif i-len(zd)==2:	
					lei2.append(n)	
				elif i-len(zd)==3:	
					lei3.append(n)	
				elif i-len(zd)==4:	
					lei4.append(n)			
			if new>1 and i>0:
				if i-len(zd)>0:
					for j in n:
						if gailv1[j]==0 or float(len(n)/(i-len(zd)))<gailv1[j]:
							gailv1[j]=float(len(n)/(i-len(zd)))				
		m+=1
	for i in lei1:
		for j in lei2:
			for k in lei3:
				
				if set(i).issubset(set(j))and len(list(set(j)-set(i)))==1:
					pan[list(set(j)-set(i))[0]]=-2
					#print list(set(j)-set(i))[0],1
					#time.sleep(100)
				if set(i).issubset(set(k)) and len(list(set(k)-set(i)))==2:
					pan[list(set(k)-set(i))[0]]=-2
					pan[list(set(k)-set(i))[1]]=-2
					#print list(set(k)-set(i))[0],list(set(k)-set(i))[1]
					#time.sleep(100)					
				if set(j).issubset(set(k))and len(list(set(k)-set(j)))==1:
					pan[list(set(k)-set(j))[0]]=-2
					#print list(set(k)-set(j))[0],3
					#time.sleep(100)

	for i in lei1:
		for j in lei1:
			if i!=j and set(i).issubset(set(j)):
				for k in list(set(j)-set(i)):
					dian.append(k)
	for i in lei2:
		for j in lei2:
			if i!=j and set(i).issubset(set(j)):
				for k in list(set(j)-set(i)):
					dian.append(k)		
	for i in lei3:
		for j in lei3:
			if i!=j and set(i).issubset(set(j)):
				for k in list(set(j)-set(i)):
					dian.append(k)						
	#print dian
	dian=list(set(dian))
	if dian==[]:
		new+=1
		if new>2:
			m=0	
			for i in pan:
				if i==-2:
					m+=1	
			if m==99:
				m=0	
				for i in pan:
					if i==-1:
						dian.append(m)	
					m+=1
			else:	
				ma=gailv1.index(max(gailv1))
				x=ma%30*18
				y=int(ma/30)*18			
				dm.moveto(x+5+1342,y+5+80)
				dm.leftclick()		
				dm.moveto(1342,40)			
				gailv1=[0]*500
				new=0
				intX=intY=0
				ret = dm.FindPic(1332,71,1886,370,"shibai.bmp","000000",0.7,0,intX,intY)
				if ret[0]!=-1:	
					time.sleep(1000)
					dm.moveto(1600,308)
					dm.leftclick()
					pan=[-1]*480
					time.sleep(2)
				
	else:
		for i in dian:
			x=i%30*18
			y=int(i/30)*18
			dm.moveto(x+5+1342,y+5+80)
			dm.leftclick()
		dm.moveto(1342,40)
		del dian[:]


new=0
dm.waitkey(65,0)

while 1:
	if dm.getkeystate(66):
		break
	over=0
	find(pan)
	jisuan(pan)


			
			
			

	



