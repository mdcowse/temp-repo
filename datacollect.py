#Requirements:
# -WMI module available at https://pypi.org/project/WMI/.
# -IIS Management Scripts and Tools component available at Web Server IIS Role installation.

import wmi, json

#Defining wmi objects in root\cimv2 namespace
c = wmi.WMI()
#Defining wmi objects in root\webAdministration namespace
w = wmi.WMI(namespace='root\webAdministration')

#Creating data structure for json export
data = {}

#Finding hostname and adding to data structure
data['hostname'] = c.Win32_ComputerSystem()[0].Name

#Finding OS Version and adding to data structure
data['oscaption'] = c.Win32_OperatingSystem()[0].Caption
data['osbuild'] = c.Win32_OperatingSystem()[0].BuildNumber

#Finding service names and adding to data structure
data['services'] = {}
for s in c.Win32_Service():
    data['services'][s.Name] = s.State

#Finding IIS app pools and adding to data structure
data['apps'] = []
for a in w.ApplicationPool():
    data['apps'].append(a.Name)

#Performing export data to json
f = open('data.json', 'w')
json.dump(data, f)
f.close()
