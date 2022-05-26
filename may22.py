#!/usr/bin/env python3
#coding:utf-8
import datetime
import xlrd
import xlwt
import re
import itertools

def chg_to_standard_freq(frequency):
	frequency=str(frequency)
	freq=frequency.upper()
	if freq in ['D','每天']:
		return "1234567"
	elif 'X' in freq:
		res=''
		for i in range(1,8):
			if str(i) not in freq:
				res+=str(i)
		return res
	else:
		return freq

def chg_to_pydatetime(time):  #将02MAY20等类似格式改为python的datetime格式，便于计算和遍历
	timeStr=time.upper()
	timeStr=timeStr.strip()
	timeStr=timeStr.replace("-","")
	if len(timeStr)==6:timeStr='0'+timeStr
	year1,mon1,day1=timeStr[5:],timeStr[2:5],timeStr[0:2]
	if len(year1)==2:year1='20'+year1
	month={"JAN":1,"FEB":2,"MAR":3,"APR":4,"MAY":5,"JUN":6,"JUL":7,"AUG":8,"SEP":9,"OCT":10,"NOV":11,"DEC":12}
	month2alp={1:'JAN',2:'FEB',3:'MAR',4:'APR',5:'MAY',6:'JUN',7:'JUL',8:'AUG',9:'SEP',10:'OCT',11:'NOV',12:'DEC'}
	t1=datetime.date(int(year1),int(month[mon1]),int(day1))
	return t1

def time_std(time):
	if len(time)<4:
		time=(4-len(time))*'0'+time
	return time

#计算时差转换后的当地时刻
def time_revise(str_t1,delta):
	hour=int(str_t1[0:2])+int(delta)
	minute=int(str_t1[-2:])
	if (delta/0.5)%2!=0:
		minute=minute+(delta-int(delta))*60
	if minute>=60:
		minute=minute-60
		hour=hour+1
	if minute<0:
		minute=60+minute
		hour=hour-1
	if hour>=24:
		hour=hour-24
	elif hour<0:
		hour=24+hour
	else:
		hour=hour
	hour=int(hour)
	minute=int(minute)
	hour=str(hour)
	minute=str(minute)
	if len(hour)==1:hour='0'+str(hour)
	if len(minute)==1:minute='0'+str(minute)
	return(hour+minute)

#计算日期是否有偏差
def date_revise(str_t1,delta):
	hour=int(str_t1[0:2])+int(delta)
	minute=int(str_t1[-2:])
	if (delta/0.5)%2!=0:
		minute=minute+(delta-int(delta))*60
	if minute>=60:
		minute=minute-60
		hour=hour+1
	if minute<0:
		minute=60+minute
		hour=hour-1
	if hour>=24:
		return 1
	elif hour<0:
		return -1
	else:
		return 0

def date_revise_2nd(*str_t):
	day_plus=0
	if int(str_t[1])<int(str_t[0]):
		day_plus+=1
	if len(str_t)==3:
		if int(str_t[2])<int(str_t[1]):
			day_plus+=1
	return day_plus

def to_datetime_type(s):
	s=s.split('-')
	return datetime.date(int(s[0]),int(s[1]),int(s[2]))

def aero_type_modify(planes):  #提取 B77W/A33L/333/359 的第一个机型，并转换乱七八糟的写法为airflite标准机型
	chg={'77W':'773','350':'359'}
	plane=planes.strip().split('/')[0][1:]
	if plane in chg:plane=chg[plane]
	return plane

def summer(seg_item):
	seg_item=list(seg_item)
	#对('MU7052', datetime.date(2022, 3, 31), 'ORD', 'PVG', '1205', '1610', 'I')进行夏令时校正
	#采用欧盟夏令时的欧洲城市
	Europe={'CDG','LHR','MXP','MAD','MAN','BEG','FRA','AMS','PRG','LUX','LGG','BUD','BRU'}
	#采用北美夏令时的美加墨城市
	NorthAmerica={'LAX','DFW','JFK','ORD','BOS','SFO','YVR','YYZ'}
	#新西兰城市
	NewZealand={'AKL'}
	#采用夏令时的澳大利亚城市
	Australia={'MEL','SYD'}
	#请务必注意，莫斯科，伊斯坦布尔，珀斯，布里斯班，夏威夷虽然属于欧美澳，但并不采用夏令时，因此不维护进集合中。
	dept_city=seg_item[2]   #当然是字符串型
	arvl_city=seg_item[3]
	dept_year=seg_item[1].year  #整型
	dept_date=seg_item[1]   #datetime型
	dept_time=seg_item[4]  #字符串型
	arvl_time=seg_item[5]
	if dept_city in Europe or arvl_city in Europe:
		#咱就是说，有没有可能，航班在欧美澳的起飞时间点，正好是落在切换冬夏令时的当天的凌晨一两点？几乎没有可能！所以把问题分析得简单一点就好。埋雷ing...
		#欧盟切换夏令时的方法为：从3月最后一个星期天到10月最后一个星期天，实行夏令时。10月最后一个星期天，就要改成冬令时。
		d_start=datetime.date(dept_year,3,31)
		while d_start.isoweekday()!=7:
			d_start-=datetime.timedelta(1)
		d_end=datetime.date(dept_year,10,31)
		while d_end.isoweekday()!=7:
			d_end-=datetime.timedelta(1)
		if dept_city in Europe:  #始发为国外，这种情况比较单纯。
			if d_start<=dept_date<d_end:
				seg_item[4]=time_revise(dept_time,1)
				seg_item[1]+=datetime.timedelta(date_revise(dept_time,1))
		else:  #落地为国外，始发为国内。这种情况较为复杂，需要根据国内日期去推断国外的始发日期（按冬令时即可，一个小时的误差无碍结论），再根据到达国外的时刻，推断到达国外的日期。最后确定一个结果：是否需要对arvl_time使用summer 函数。
			d=dept_date
			dept_foreign_time=time_revise(dept_time,time_zone_offset[arvl_city])
			dept_foreign_date_offset=date_revise(dept_time,time_zone_offset[arvl_city])
			if int(arvl_time)<int(dept_foreign_time):dept_foreign_date_offset+=1
			d+=datetime.timedelta(dept_foreign_date_offset)  #终于得到了按照国外当地时计算的国外落地日期，以此计算夏令时是否需要转换，才是对的。
			if d_start<=d<d_end:
				seg_item[5]=time_revise(arvl_time,1)
#		if dept_city=='MXP' or arvl_city=='MXP':print(seg_item)
		return tuple(seg_item)


	#北美的夏令时切换规则为：从每年3月的第2个周日开始，到每年11月的第一个周日结束。夏威夷不使用夏令时。亚利桑那不使用夏令时。
	#务必注意边界问题，因为已经发现了凌晨1点多起飞的航班。它在夏令时切换日的转换，很重要。是软件品质的关键。多次测试不同场景。
	elif dept_city in NorthAmerica or arvl_city in NorthAmerica:
#		count=0
		d_start=datetime.date(dept_year,3,1)
		while d_start.isoweekday()!=7:
			d_start+=datetime.timedelta(1)
		d_start+=datetime.timedelta(1)
		while d_start.isoweekday()!=7:
			d_start+=datetime.timedelta(1)

		d_end=datetime.date(dept_year,11,1)
		while d_end.isoweekday()!=7:
			d_end+=datetime.timedelta(1)

		if dept_city in NorthAmerica:  #始发为国外，这种情况比较单纯。
			if d_start<=dept_date<d_end:
				seg_item[4]=time_revise(dept_time,1)
				seg_item[1]+=datetime.timedelta(date_revise(dept_time,1))
		else:  #落地为国外，始发为国内。这种情况较为复杂，需要根据国内日期去推断国外的始发日期（按冬令时即可，一个小时的误差无碍结论），再根据到达国外的时刻，推断到达国外的日期。最后确定一个结果：是否需要对arvl_time使用summer 函数。
			d=dept_date
			dept_foreign_time=time_revise(dept_time,time_zone_offset[arvl_city])
			dept_foreign_date_offset=date_revise(dept_time,time_zone_offset[arvl_city])
			if int(arvl_time)<int(dept_foreign_time):dept_foreign_date_offset+=1
			d+=datetime.timedelta(dept_foreign_date_offset)  #终于得到了按照国外当地时计算的国外落地日期，以此计算夏令时是否需要转换，才是对的。
			if d_start<=d<d_end:
				seg_item[5]=time_revise(arvl_time,1)

		return tuple(seg_item)

	#新西兰是9月最后一个星期日到第二年4月第一个星期日。
	elif dept_city in NewZealand or arvl_city in NewZealand:
		#注意，这里start和end的含义与北半球不同
		d_start=datetime.date(dept_year,9,30)
		while d_start.isoweekday()!=7:
			d_start-=datetime.timedelta(1)

		d_end=datetime.date(dept_year,4,1)
		while d_end.isoweekday()!=7:
			d_end+=datetime.timedelta(1)


		if dept_city in NewZealand:  #始发为国外，这种情况比较单纯。
			if dept_date>=d_start or dept_date<d_end:
				seg_item[4]=time_revise(dept_time,1)
				seg_item[1]+=datetime.timedelta(date_revise(dept_time,1))
		else:  #落地为国外，始发为国内。这种情况较为复杂，需要根据国内日期去推断国外的始发日期（按冬令时即可，一个小时的误差无碍结论），再根据到达国外的时刻，推断到达国外的日期。最后确定一个结果：是否需要对arvl_time使用summer 函数。
			d=dept_date
			dept_foreign_time=time_revise(dept_time,time_zone_offset[arvl_city])
			dept_foreign_date_offset=date_revise(dept_time,time_zone_offset[arvl_city])
			if int(arvl_time)<int(dept_foreign_time):dept_foreign_date_offset+=1
			d+=datetime.timedelta(dept_foreign_date_offset)  #终于得到了按照国外当地时计算的国外落地日期，以此计算夏令时是否需要转换，才是对的。
			if d>=d_start or d<d_end:
				seg_item[5]=time_revise(arvl_time,1)

		return tuple(seg_item)
	

	#澳大利亚是10月第一个周日到第二年4月第一个周日。并非澳大利亚所有地区都是用夏令时 (DST)。使用夏令时的州包括：澳大利亚首都特区、新南威尔士、南澳大利亚、塔斯马尼亚、维多利亚。而剩余的西澳大利亚、昆士兰、北领地以及尤克拉不实行夏令时。昆士兰州包括东航站点BNE。
	elif dept_city in Australia or arvl_city in Australia:
		#注意，这里start和end的含义与北半球不同
		d_start=datetime.date(dept_year,10,1)
		while d_start.isoweekday()!=7:
			d_start+=datetime.timedelta(1)

		d_end=datetime.date(dept_year,4,1)
		while d_end.isoweekday()!=7:
			d_end+=datetime.timedelta(1)

		if dept_city in Australia:  #始发为国外，这种情况比较单纯。
			if dept_date>=d_start or dept_date<d_end:
				seg_item[4]=time_revise(dept_time,1)
				seg_item[1]+=datetime.timedelta(date_revise(dept_time,1))
		else:  #落地为国外，始发为国内。这种情况较为复杂，需要根据国内日期去推断国外的始发日期（按冬令时即可，一个小时的误差无碍结论），再根据到达国外的时刻，推断到达国外的日期。最后确定一个结果：是否需要对arvl_time使用summer 函数。
			d=dept_date
			dept_foreign_time=time_revise(dept_time,time_zone_offset[arvl_city])
			dept_foreign_date_offset=date_revise(dept_time,time_zone_offset[arvl_city])
			if int(arvl_time)<int(dept_foreign_time):dept_foreign_date_offset+=1
			d+=datetime.timedelta(dept_foreign_date_offset)  #终于得到了按照国外当地时计算的国外落地日期，以此计算夏令时是否需要转换，才是对的。
			if d>=d_start or d<d_end:
				seg_item[5]=time_revise(arvl_time,1)

		return tuple(seg_item)
				
	else:
		#出发或到达非欧美澳，完全不用校正，直接原样返回seg
		return tuple(seg_item)



if __name__=='__main__':
	dic_month_reverse={'Jan':1,'Feb':2,'Mar':3,'Apr':4,'May':5,'Jun':6,'Jul':7,'Aug':8,'Sep':9,'Oct':10,'Nov':11,'Dec':12}
	dic_month={1:'JAN',2:'FEB',3:'MAR',4:'APR',5:'MAY',6:'JUN',7:'JUL',8:'AUG',9:'SEP',10:'OCT',11:'NOV',12:'DEC'}
	domestic_code3=["AAT","ACX","AEB","AKA","AKU","AOG","AQG","AVA","AYN","BAR","BAV","BFJ","BFU","BHY","BPE","BPX","BSD",
	"LHW","CAN","CDE","CGD","CGO","CGQ","CHG","CIF","CIH","CKG","CNI","CTU","CWJ","CZX","DAT","DAX","DBC","DCY","DDG","DIG",
	"DLC","DLU","DNH","DOY","DQA","DSN","DTU","DYG","DZU","ENH","ENY","ERL","FNJ","FOC","FUD","FUG","FUO","FYN","GHN","GMQ",
	"GOQ","GYS","GYU","HAK","HBQ","HDG","HEK","HET","HFE","HGH","CSX","HIA","HJJ","HKF","HLD","HLH","HMI","HNY","HPG","HRB",
	"HSN","HTN","HTT","HUZ","HXD","HYN","HZG","HZH","INC","IQM","IQN","JDZ","JGD","JGN","JGS","JHG","JIC","JIL","JIQ","JIU",
	"JJN","JMJ","JMU","JNG","JNZ","JUH","JUZ","JXA","JZH","KCA","KGT","KHG","KHN","KJH","KJH","KJI","KMG","KNC","KOW","KRL",
	"KRY","KWE","KWL","LCX","LDG","LDS","LDX","LEI","LFQ","LHK","LIA","LJG","LLB","LLF","LLV","LNJ","LNL","LPF","LUM","LUZ",
	"LXA","LXI","LYA","LYG","LYI","LZH","LZO","LZY","MDG","MIG","MLN","MXZ","NAO","NAY","NBS","NDG","NGB","NGQ","NKG","NLH",
	"NLT","NNG","NNY","NTG","NZH","OHE","PEK","PKX","PNJ","PVG","PZI","QSZ","RGA","RIZ","RKZ","RLK","SCM","SHA","SHE","SHF",
	"SHS","SHU","SJW","SQD","SSF","SWA","SXJ","SYM","SYX","SZX","TAO","TCG","TCZ","TEN","TGO","THE","THQ","TLQ","TLQ","TNA",
	"TNH","TSN","TVS","TXN","TYN","UCB","ULN","URC","UYN","WDS","WEF","WEH","WGN","WMT","WNH","WNZ","WUA","WUH","WUS","WUT",
	"WUX","WUZ","WXN","XAI","XEN","XFN","XIC","XIL","XIN","XIY","XMN","XNN","XNT","XUZ","XYI","YBP","YCU","YIC","YIE","YIH",
	"YIN","YIW","CSX","YKH","YLN","YNJ","YNT","YNZ","YSQ","YTY","YUA","YUS","YZY","ZAT","ZBX","ZDC","ZGC","ZHA","ZHY","ZQZ",
	"ZSJ","ZSX","ZUH","ZYI","ZAZ","HKG"]

	#注意，请务必维护标准时差（或称冬令时），因为后面我们会通过夏令时自动检测模块，来校正欧美澳的夏令时问题
	time_zone_offset={'CDG':-7,'LHR':-8,'MXP':-7,'LAX':-16,'PER':0,'MNL':0,'BKK':-1,'CEB':0,'CGK':-1,'DFW':-13,'DEL':-2.5,
	'DXB':-4,'ICN':1,'JFK':-13,'KIX':1,'MAD':-7,'MEL':2,'ORD':-14,'BOS':-13,'PNH':-1,'SFO':-16,'SGN':-1,'SIN':0,'SYD':2,
	'TPE':0,'YVR':-16,'YYZ':-13,'MAN':-8,'VVO':2,'NRT':1,'SVO':-5,'KHI':-3,'BEG':-7,'KUL':0,'FRA':-7,'CTS':1,'AMS':-7,
	'CMB':-2.5,'FUK':1,'CCU':-2.5,'HAN':-1,'PRG':-7,'BOM':-2.5,'LUX':-7,'LGG':-7,'BUD':-7,'DAC':-2,'BRU':-7,'AKL':4}

	jihua_set=set()
	data_jihua=xlrd.open_workbook('计划表.xls')
	table_jihua=data_jihua.sheets()[0]
	rows_num1=table_jihua.nrows
	cols_num1=table_jihua.ncols
	for i in range(rows_num1):
		line=table_jihua.row_values(i)
		service_type=line[0]
		if line[3] in ['','航线']:continue
		flight_num=line[2]
		flight_num_list=[]

		#航班号
		if flight_num.count('/')!=0:          #两个航班号
			flight_num0=flight_num[:flight_num.index('/')]   #MU2721
			delta=len(flight_num)-flight_num.index('/')-1
			flight_num1=flight_num[:flight_num.index('/')][:-delta]+flight_num[-delta:]   #MU2722
			flight_num_list.append(flight_num0)
			flight_num_list.append(flight_num1)
		else:
			flight_num_list.append(flight_num)

		#日期
		date_range=line[4]
		date_list=[]

		#班期
		freq=line[5]

		#机型
		aero_type=aero_type_modify(line[6])

		#departure,arrival
		dp,ar=line[7].strip(),line[11].strip()
		dp2,ar2=line[11].strip(),line[16].strip()

		dp_time,ar_time=line[8],line[9]
		dp_time,ar_time=time_std(dp_time),time_std(ar_time)
		dp_time2,ar_time2=line[12],line[14]
		dp_time2,ar_time2=time_std(dp_time2),time_std(ar_time2)

		#航段
		seg_list=[]
		seg_list.append((dp,ar))
		if ar2!='':seg_list.append((dp2,ar2))

		if dp not in domestic_code3:
			dp_time=time_revise(dp_time,time_zone_offset[dp])
		if ar not in domestic_code3:
			ar_time=time_revise(ar_time,time_zone_offset[ar])
		if dp2 not in domestic_code3 and dp2!='':
			dp_time2=time_revise(dp_time2,time_zone_offset[dp2])
		if ar2 not in domestic_code3 and ar2!='':
			ar_time2=time_revise(ar_time2,time_zone_offset[ar2])

		try:
			date_range=date_range.strip()
			digits=re.split(r'[^0-9/]+',date_range)  #digits=['2022', '5', '8', '10', '29', '5', '9/10/13', '']
			digits.pop()  #最后一位必定为空值，踢掉吧
			if len(digits)==3:  #最简单的场景，一个日期或者几个日期
				dy=digits[-1].split('/')
				mth=digits[1]
				yr=digits[0]
				for d in dy:
					date_list.append(datetime.date(int(yr),int(mth),int(d)))

			elif len(digits)==5:  #一段日期，也就是2022年5月1日至7月21日
				date_from=datetime.date(int(digits[0]),int(digits[1]),int(digits[2]))
				date_to=datetime.date(int(digits[0]),int(digits[3]),int(digits[4]))
				dy=date_from
				while dy<=date_to:
					if str(dy.weekday()+1) in chg_to_standard_freq(freq):
						date_list.append(dy)
					dy=dy+datetime.timedelta(1)				

			elif len(digits)==6:  #需要注意，可能是2022年12月1日至2023年2月1日
				date_from=datetime.date(int(digits[0]),int(digits[1]),int(digits[2]))
				date_to=datetime.date(int(digits[3]),int(digits[4]),int(digits[5]))
				dy=date_from
				while dy<=date_to:
					if str(dy.weekday()+1) in chg_to_standard_freq(freq):
						date_list.append(dy)
					dy=dy+datetime.timedelta(1)		

			elif len(digits)==7:  #肯定是备注了需要剔除的日期
				date_from=datetime.date(int(digits[0]),int(digits[1]),int(digits[2]))
				date_to=datetime.date(int(digits[0]),int(digits[3]),int(digits[4]))
				dy=date_from
				while dy<=date_to:
					if str(dy.weekday()+1) in chg_to_standard_freq(freq):
						date_list.append(dy)
					dy=dy+datetime.timedelta(1)	
				delete_dy=digits[-1].split('/')
				delete_mth=digits[-2]
				for delete_d in delete_dy:
					delete_date=datetime.date(int(digits[0]),int(delete_mth),int(delete_d))
					if delete_date in date_list:
						date_list.remove(delete_date)
			else:
				print('无法理解 '+date_range+' Orz')
		except Exception as e:
			print(e.args)
			print(line,date_range,"格式不符")


		if len(flight_num_list)==1:#只一个航班号，分为单程和联程两种情况
			if dp not in domestic_code3: #国际始发，做一个预处理
				for k,each_date in enumerate(date_list):
					date_list[k]+=datetime.timedelta(date_revise(time_std(line[8]),time_zone_offset[dp]))
			for item in date_list:#如果只有一个航段，国际始发或国内始发都一样
				information=(flight_num_list[0],item,dp,ar,dp_time,ar_time,service_type,aero_type)
				jihua_set.add(information)
			if len(seg_list)==2:  #如果是联程，
				for k,each_date in enumerate(date_list):#飞行过程中的跨天分析
					date_list[k]+=datetime.timedelta(date_revise_2nd(time_std(line[8]),time_std(line[12])))
				if dp not in domestic_code3:  #始发为国际，则中间点必为国内。把日期差倒回去。
					date_list[k]+=datetime.timedelta(0-date_revise(time_std(line[8]),time_zone_offset[dp]))
				for item in date_list:
					information=(flight_num_list[0],item,dp2,ar2,dp_time2,ar_time2,service_type,aero_type)
					jihua_set.add(information)			

		elif len(flight_num_list)==2:#两个航班号合并书写，只能是单程往返这一种情况。包括DDD,DID
			for item in date_list:#DD,DI
				information=(flight_num_list[0],item,dp,ar,dp_time,ar_time,service_type,aero_type)
				jihua_set.add(information)
			if dp2 not in domestic_code3: 
				for k,each_date in enumerate(date_list):
					date_list[k]+=datetime.timedelta(date_revise(time_std(line[12]),time_zone_offset[dp2])+date_revise_2nd(time_std(line[8]),time_std(line[12])))
			else:
				for k,each_date in enumerate(date_list):
					date_list[k]+=datetime.timedelta(date_revise_2nd(time_std(line[8]),time_std(line[12])))

			for item in date_list:
				information=(flight_num_list[1],item,dp2,ar2,dp_time2,ar_time2,service_type,aero_type)
				jihua_set.add(information)

	jihua_set_summer_fixed=set()
	for item in jihua_set:
		item_summer_fixed=summer(item)
		jihua_set_summer_fixed.add(item_summer_fixed)
#		print(item_summer_fixed)

	set_airflite=set()
	f=open('AF.txt','r+',encoding='utf-8')
	for line in itertools.islice(f,1,None):
		line=line.strip()
		if len(line)!=0:  #空行不用考虑
			line=line.split('\t')
			svc_type=line[12]
			if svc_type in ['C','G','P','O','H','Y','I']:  #service type非cgpoh 不用考虑
				line[0]=line[0].replace(" ","")
				flight_number=line[0]
				dept,arvl=line[4],line[6]
				dept_time,arvl_time=line[5].replace(":",""),line[7].replace(":","")
				plane_type=line[8].strip().split(' ')[-1]  #789
				if len(dept_time)<4:dept_time=time_std(dept_time)
				if len(arvl_time)<4:arvl_time=time_std(arvl_time)


				date_from=chg_to_pydatetime(line[1])
				date_to=chg_to_pydatetime(line[2])
				flt_freq=chg_to_standard_freq(line[3])
				day=date_from
				while day<=date_to:
					if str(day.weekday()+1) in flt_freq:
						set_airflite.add((flight_number,day,dept,arvl,dept_time,arvl_time,svc_type,plane_type))
					day=day+datetime.timedelta(1)

	f.close()
#{('MU7188', datetime.date(2021, 7, 22), 'MXP', 'TNA', '1630', '0855','Y'), ('MU7587', datetime.date(2021, 10, 16), 
	
	jihua_dic={}
#	dup_jiha_dic={}
	for item in sorted(jihua_set_summer_fixed):
		if item[0]+' / '+str(item[1]) not in jihua_dic:
			jihua_dic[item[0]+' / '+str(item[1])]=[(item[2],item[3],item[4],item[5],item[6],item[7])]
		else:
			print(item[0]+' / '+str(item[1])+' '+item[2]+' '+item[3]+' '+item[4]+' '+item[5]+' '+item[6],' 重复录入')
#			dup_jiha_dic[item[0]+' / '+str(item[1])]=(item[2],item[3],item[4],item[5],item[6])
			jihua_dic[item[0]+' / '+str(item[1])].append((item[2],item[3],item[4],item[5],item[6],item[7]))
	dic_airflite={}
	for item in set_airflite:
		if item[0]+' / '+str(item[1]) not in dic_airflite:
			dic_airflite[item[0]+' / '+str(item[1])]=(item[2],item[3],item[4],item[5],item[6],item[7])
		elif dic_airflite[item[0]+' / '+str(item[1])][4]=='I':  #如果有非I的重复航班，则替换掉现有的
			dic_airflite[item[0]+' / '+str(item[1])]=(item[2],item[3],item[4],item[5],item[6],item[7])
		else:
			continue

	res_lst=[]
	for i in jihua_dic:
		res_lst.append(i)
	res_lst.sort(key=lambda x:x)
	#print(res_lst)
	#['MU7523 / 2022-05-01', 'MU7523 / 2022-05-02',...]

	airflite_lst=[]
	for i in dic_airflite:
		airflite_lst.append(i)
	airflite_lst.sort(key=lambda x:x)
#	print(res_lst)
	print('计划表','\t','Airflite')
	duplicated_set=set()

	wb=xlwt.Workbook()
	sheet1=wb.add_sheet('Sheet1',cell_overwrite_ok=True)
	style0 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on')
	style1 = xlwt.easyxf('font: name Arial, color-index black, bold on;align: wrap on, vert center, horiz center;')

	sheet1.write(0,0,"计划表")
	sheet1.write(0,8,"Airflite")
	#print(jihua_dic)
	#{'MU7523 / 2022-05-01': [('PVG', 'NRT', '0845', '1230', 'Y', '773')], 'MU7523 / 2022-05-02': [('PVG', 'NRT', '0845', '1230', 'Y', '773')], 
	row=1
	for i in res_lst:
		if i in dic_airflite:
			for k in jihua_dic[i]:
				if (k[0],k[1],k[2],k[3],k[4],k[5])!=(dic_airflite[i][0],dic_airflite[i][1],dic_airflite[i][2],dic_airflite[i][3],dic_airflite[i][4],dic_airflite[i][5]):
					
					sheet1.write(row,0,str(i).split('/')[0])
					sheet1.write(row,8,str(i).split('/')[0])

					sheet1.write(row,1,str(i).split('/')[1])
					sheet1.write(row,9,str(i).split('/')[1])

					if (k[0]+'-'+k[1])!=(dic_airflite[i][0]+'-'+dic_airflite[i][1]):
						sheet1.write(row,2,k[0]+'-'+k[1],style0)
						sheet1.write(row,10,dic_airflite[i][0]+'-'+dic_airflite[i][1],style0)
					else:
						sheet1.write(row,2,k[0]+'-'+k[1])
						sheet1.write(row,10,dic_airflite[i][0]+'-'+dic_airflite[i][1])	

					if (str(to_datetime_type(i[-10:]).weekday()+1))!=(str(to_datetime_type(i[-10:]).weekday()+1)):
						sheet1.write(row,3,str(to_datetime_type(i[-10:]).weekday()+1),style0)
						sheet1.write(row,11,str(to_datetime_type(i[-10:]).weekday()+1),style0)
					else:
						sheet1.write(row,3,str(to_datetime_type(i[-10:]).weekday()+1))
						sheet1.write(row,11,str(to_datetime_type(i[-10:]).weekday()+1))		

					if (k[2]+'-'+k[3])!=(dic_airflite[i][2]+'-'+dic_airflite[i][3]):
						sheet1.write(row,4,k[2]+'-'+k[3],style0)
						sheet1.write(row,12,dic_airflite[i][2]+'-'+dic_airflite[i][3],style0)
					else:
						sheet1.write(row,4,k[2]+'-'+k[3])
						sheet1.write(row,12,dic_airflite[i][2]+'-'+dic_airflite[i][3])		

					if (k[4])!=(dic_airflite[i][4]):
						sheet1.write(row,5,k[4],style0)
						sheet1.write(row,13,dic_airflite[i][4],style0)
					else:
						sheet1.write(row,5,k[4])
						sheet1.write(row,13,dic_airflite[i][4])

					if (k[5])!=(dic_airflite[i][5]):
						sheet1.write(row,6,k[5],style0)
						sheet1.write(row,14,dic_airflite[i][5],style0)
					else:
						sheet1.write(row,6,k[5])
						sheet1.write(row,14,dic_airflite[i][5])				

					sheet1.write(row,7,"—")											
					row+=1
				duplicated_set.add(i)
		else:
			for j in jihua_dic[i]:
				sheet1.write(row,0,str(i).split('/')[0])
				sheet1.write(row,1,str(i).split('/')[1])
				sheet1.write(row,2,j[0]+'-'+j[1])
				sheet1.write(row,3,str(to_datetime_type(i[-10:]).weekday()+1))
				sheet1.write(row,4,j[2]+'-'+j[3])
				sheet1.write(row,5,j[4])
				sheet1.write(row,6,j[5])
				sheet1.write(row,7,"—")
				row+=1
				duplicated_set.add(i)
	for i in airflite_lst:
		if i not in duplicated_set:
			sheet1.write(row,8,str(i).split('/')[0])
			sheet1.write(row,9,str(i).split('/')[1])
			sheet1.write(row,10,dic_airflite[i][0]+'-'+dic_airflite[i][1])
			sheet1.write(row,11,str(to_datetime_type(i[-10:]).weekday()+1))
			sheet1.write(row,12,dic_airflite[i][2]+'-'+dic_airflite[i][3])
			sheet1.write(row,13,dic_airflite[i][4])
			sheet1.write(row,14,dic_airflite[i][5])
			sheet1.write(row,7,"—")
#			row+=1
#			sheet1.write(row,1,str(i)+' '+dic_airflite[i][0]+'-'+dic_airflite[i][1]+' '+str(to_datetime_type(i[-10:]).weekday()+1)+' '+dic_airflite[i][2]+'-'+dic_airflite[i][3]+' '+dic_airflite[i][4])
			row+=1
	sheet1.write_merge(0,0,0,6,'计划表',style1)
	sheet1.write_merge(0,0,8,14,'Airflite',style1)
	file_name=str(datetime.datetime.today().strftime('%Y-%m-%d-%H：%M'))+'货运计划比对.xls'
	wb.save(file_name)
