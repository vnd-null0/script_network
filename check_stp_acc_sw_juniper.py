# $language = "python"
# $interface = "1.0"

sHeader = 'sw_ip >>> stp, stp_region\n'
	
juniper_command_list = ["show spanning-tree bridge | no-more",
						"show spanning-tree mstp configuration | no-more"]

SCRIPT_TAB = crt.GetScriptTab()
SCRIPT_TAB.Screen.Synchronous = True
SCRIPT_TAB.Screen.IgnoreEscape = True

def CaptureOutputOfCommand(command, prompt):
	if not crt.Session.Connected:
		return "[ERROR: Not Connected.]"
	SCRIPT_TAB.Screen.Send(command + '\r')
	SCRIPT_TAB.Screen.WaitForString('\r')
	return SCRIPT_TAB.Screen.ReadString(prompt)
	
def process_juniper_sw(sw_ip, username='duyvn', password ='yuDtahN@2091993'):
	crt.Screen.Send("\r")
	crt.Screen.WaitForString("$")
	crt.Screen.Send("ssh %s@%s\n"  %(username,sw_ip))
	while True :
		result = crt.Screen.WaitForStrings(["?","word:"], 10)
		if result == 1:
			crt.Screen.Send("yes\n")
		if result == 2:
			crt.Screen.Send("%s\n" %(password))
			break
	crt.Screen.WaitForString(">")
	data = ''
	for command in juniper_command_list:
		data += CaptureOutputOfCommand(command,">")
	crt.Screen.Send("quit\r")
	#crt.Dialog.MessageBox(data)
	#open(r"D:\output_sw_list.txt","w").write(data.replace('\r\r\n', '\r\n'))
	stp    = "OK" if 'Root port' in data else "RE_CHECK"
	stp_region = "OK" if 'VNG' in data else "RE_CHECK"
	
	return (stp, stp_region)
	
def main():
	finalData = sHeader
	for line in open(r"D:\PROGRAMING\_code_4_job_\sw_list_stp.txt"):
		sw_data = process_juniper_sw(line)
		finalData = finalData + line.rstrip('\n') + ' >>> ' + ' '.join([str(x) for x in sw_data]) + '\n'
	open(r"D:\PROGRAMING\_code_4_job_\out_sw_list_stp.txt", "w").write(finalData)
		#crt.Dialog.MessageBox('\n'.join([ str(xx) for xx in x ]))

main()