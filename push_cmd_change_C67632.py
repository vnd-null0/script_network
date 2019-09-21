# $language = "python"
# $interface = "1.0"

def init_session_juniper(ip_devices, username='xxx',password ='xxx'):
    crt.Screen.Send("\r")
    crt.Screen.WaitForString("$")
    crt.Screen.Send("ssh %s@%s\n"  %(username,ip_devices))
    while True:
        result = crt.Screen.WaitForStrings(["?","word:"], 10)
        if result == 1:
            crt.Screen.Send("yes\n")
        if result == 2:
            crt.Screen.Send("%s\n" %(password))
            break
    crt.Screen.WaitForString(">")
    crt.Screen.Send("configure\n")
    crt.Screen.WaitForString("#", 2)

def main():
    for line in open(r"D:\PROGRAMING\_code_4_job_\core_log_prefix.txt"):
        params = line.split()
        init_session_juniper(params[0])
        #crt.Screen.Send("configure\n")
        #crt.Screen.WaitForString("#")
        cmd="set system syslog host 10.73.2.240 log-prefix "+params[1]+"_"+params[2]+"_"+params[0]+'\n'\
        +"set system syslog host 10.73.2.8 log-prefix "+params[1]+"_"+params[2]+"_"+params[0]+'\n'
        crt.Screen.Send(cmd)
		# Cause a 3-second pause between sends by waiting for something "unexpected"
		# with a timeout value.
        crt.Screen.WaitForString("something_unexpected", 2)
        crt.Screen.Send("commit and-quit\n")
        crt.Screen.WaitForString("something_unexpected", 7)
        #crt.Screen.Send("quit\n")
        crt.Screen.WaitForString("something_unexpected", 2)
        crt.Screen.Send("quit\n")

main()