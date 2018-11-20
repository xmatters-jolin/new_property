# encoding: utf-8
"""Create and maintain Four Seasons entities
    
    This is a command line utility that takes a specifically formatted spreadsheet
    as input and creats the associated xMatters objects
    
    Example:
    Arguments are described via the -H command
    Here are some examples::
    
    $ python3 new_property.py -vv -c -d defaults.json all
    $ python3 new_property.py -vvv -c -d 4s.defaults.json sites

    .. _Google Python Style Guide:
    http://google.github.io/styleguide/pyguide.html
    
    """

import sys

import config
import cli

__all__ = []
__version__ = config.VERSION
__date__ = config.DATE
__updated__ = config.UPDATED

def main(argv=None):
    """ Begins the New Properties process """
    
    args = cli.process_command_line(argv, __doc__)
    args.func(args)

if __name__ == "__main__":
    if config.DEBUG:
        sys.argv.append("-h")
        sys.argv.append("-v")
    if config.TESTRUN:
        import doctest
        doctest.testmod()
    if config.PROFILE:
        import cProfile
        import pstats
        profile_filename = 'NewProperites_profile.txt'
        cProfile.run('main()', profile_filename)
        statsfile = open("profile_stats.txt", "wb")
        p = pstats.Stats(profile_filename, stream=statsfile)
        stats = p.strip_dirs().sort_stats('cumulative')
        stats.print_stats()
        statsfile.close()
        sys.exit(0)
    sys.exit(main())



today = (time.strftime("%Y-%m-%d"))
print(today)

base_URL = 'https://capitalgroup-np.hosted.xmatters.com/api/xm/1'

workbook = xlsxwriter.Workbook('GroupReportCapGroupNP.xlsx')
worksheet = workbook.add_worksheet()
worksheet.write('A1' , 'group_name')
worksheet.write('B1' , 'group_description')
worksheet.write('C1'  , 'user_name')
worksheet.write('D1'  , 'shift')
worksheet.write('E1'  , 'position')
worksheet.write('F1'  , 'delay')
worksheet.write('G1'  , 'first_name')
worksheet.write('H1'  , 'last_name')
worksheet.write('I1' , 'DRE')
worksheet.write('J1', 'android_phone')
worksheet.write('K1', 'android_tablet')
worksheet.write('L1', 'home_email')
worksheet.write('M1', 'home_phone')
worksheet.write('N1', 'ipad')
worksheet.write('O1', 'iphone')
worksheet.write('P1', 'mobile_phone')
worksheet.write('Q1', 'other_phone')
worksheet.write('R1', 'sms_phone')
worksheet.write('S1' , 'work_email')
worksheet.write('T1' , 'work_phone')
worksheet.write('U1' , 'Q10')
worksheet.write('V1' , 'Z10')
worksheet.write('W1' , 'Z30')

group_name = 'None'
group_description = 'None'
user_name = 'None'
shift = 'None'
position = 'None'
delay = 'None'
first_name = 'None'
last_name = 'None'
DRE = 'None'
android_phone = 'None'
android_tablet = 'None'
home_email = 'None'
home_phone = 'None'
ipad = 'None'
iphone = 'None'
mobile_phone = 'None'
other_phone = 'None'
personal_phone = 'None'
sms_phone = 'None'
work_email = 'None'
work_phone = 'None'


headers = {
			'content-type': "application/json",
			'authorization': "Basic eG0tam9saW46QEw0bl9UdXJpbjk=",
			'cache-control': "no-cache"
		}
aurl = base_URL + '/groups?offset=0&limit=1000'
aresponse = requests.get(aurl, headers=headers)
#aresponse = requests.get(aurl, auth=HTTPBasicAuth('xm-jolin', 'ParkCity14!'))
print(aurl)
print(aresponse)
ajson = aresponse.json()
i = 2
g = 1
for a in ajson['data']:
	g = g+1
	#print(g)
	group = a['id']
	group_name = a['targetName']
	group_description = a['description']
	print(group_name)

	#gurl = base_URL + '/groups/' + group + '/members?embed=shifts'
	gurl = base_URL + '/on-call?groups=' + group + '&embed=shift,members.owner&from=' + today + 'T08:00:00Z&to=' + today + 'T23:59:59Z'
	#gresponse = requests.get(gurl, auth=HTTPBasicAuth('kapgar', 'ParkCity14!'))
	gresponse = requests.get(gurl, headers=headers)
	print(gurl)
	print(gresponse)

	if (gresponse.status_code == 200):
		gjson = gresponse.json();

		for d in gjson['data']:
			if 'shift' in d:
				shift = d['shift']['name']
				#print(shift)
				for m in d['members']['data']:
					position = m['position']
					delay = m['delay']
					user_name = m['member']['targetName']
					uurl = base_URL + '/people/' + user_name + '?embed=roles'
					#uresponse = requests.get(uurl, auth=HTTPBasicAuth('kapgar', 'ParkCity14!'))
					uresponse = requests.get(uurl, headers=headers)
					if (uresponse.status_code == 200):
						ujson = uresponse.json()
						first_name = ujson['firstName']
						#print (first_name)
						last_name = ujson['lastName']
						if 'properties' not in ujson:
							#print('No Properties')
							DRE = 'None'
						else:
							#print(ujson['properties'])
							if '2016 DRE' in ujson['properties']:
								DRE = ujson['properties']['2016 DRE']
						android_phone = 'None'
						android_tablet = 'None'
						home_email = 'None'
						home_phone = 'None'
						ipad = 'None'
						iphone = 'None'
						mobile_phone = 'None'
						other_phone = 'None'
						personal_phone = 'None'
						sms_phone = 'None'
						work_email = 'None'
						work_phone = 'None'
						q10 = 'None'
						z10 = 'None'
						z30 = 'None'
						durl = base_URL + '/people/' + user_name + '/devices'
						#dresponse = requests.get(durl, auth=HTTPBasicAuth('kapgar', 'ParkCity14!'))
						dresponse = requests.get(durl, headers=headers)
						if (dresponse.status_code == 200):
							print(durl)
							print(dresponse)
							djson = dresponse.json()
							for l in djson['data']:
								name = l['name']
								if name == 'Android phone':
									android_phone = l['description']
								elif name == 'Android tablet':
									android_tablet = l['description']
								elif name == 'Home Email':
									home_email = l['emailAddress']
								elif name == 'Home Phone':
									home_phone = l['phoneNumber']
								elif name == 'iPad':
									ipad = l['description']
								elif name == 'iPhone':
									iphone = l['description']
								elif name == 'Mobile Phone':
									mobile_phone = l['phoneNumber']
								elif name == 'Other Phone':
									other_phone = l['phoneNumber']
								elif name == 'SMS Phone':
									sms_phone = l['phoneNumber']
								elif name == 'Work Email':
									work_email = l['emailAddress']
								elif name == 'Work Phone':
									work_phone = l['phoneNumber']
								elif name == 'Q10':
									q10 = l['phoneNumber']
								elif name == 'Z10':
									z10 = l['phoneNumber']
								elif name == 'Z30':
									z30 = l['phoneNumber']
						print('Updating Worksheet - ' + shift)
						worksheet.write('A' + str(i), group_name)
						worksheet.write('B' + str(i), group_description)
						worksheet.write('C' + str(i), user_name)
						worksheet.write('D' + str(i), shift)
						worksheet.write('E' + str(i), position)
						worksheet.write('F' + str(i), delay)
						worksheet.write('G' + str(i), first_name)
						worksheet.write('H' + str(i), last_name)
						worksheet.write('I' + str(i), DRE)
						worksheet.write('J' + str(i), android_phone)
						worksheet.write('K' + str(i), android_tablet)
						worksheet.write('L' + str(i), home_email)
						worksheet.write('M' + str(i), home_phone)
						worksheet.write('N' + str(i), ipad)
						worksheet.write('O' + str(i), iphone)
						worksheet.write('P' + str(i), mobile_phone)
						worksheet.write('Q' + str(i), other_phone)
						worksheet.write('R' + str(i), sms_phone)
						worksheet.write('S' + str(i), work_email)
						worksheet.write('T' + str(i), work_phone)
						worksheet.write('U' + str(i), q10)
						worksheet.write('V' + str(i), z10)
						worksheet.write('W' + str(i), z30)
						i = i + 1
			else:
				print('No Shifts are defined.')
				worksheet.write('A' + str(i), group_name)
				worksheet.write('B' + str(i), group_description)
				worksheet.write('C' + str(i), '')
				worksheet.write('D' + str(i), 'No Shifts defined')
				i = i + 1

workbook.close()





