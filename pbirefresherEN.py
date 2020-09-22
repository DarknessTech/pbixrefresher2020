import time
import os
import sys
import argparse
import psutil
import pyautogui
from pywinauto.application import Application
from pywinauto import timings

def main():   
	# Parse arguments from cmd
	parser = argparse.ArgumentParser()
	parser.add_argument("workbook", help = "Path to .pbix file")
	parser.add_argument("--workspace", help = "name of online Power BI service work space to publish in", default = "My workspace")
	parser.add_argument("--refresh-timeout", help = "refresh timeout", default = 30000, type = int)
	parser.add_argument("--no-publish", dest='publish', help="don't publish, just save", default = True, action = 'store_false' )
	parser.add_argument("--init-wait", help = "initial wait time on startup", default = 60, type = int)
	args = parser.parse_args()

	timings.after_clickinput_wait = 1
	WORKBOOK = args.workbook
	WORKSPACE = args.workspace
	INIT_WAIT = args.init_wait
	REFRESH_TIMEOUT = args.refresh_timeout

	# Kill running PBI
	PROCNAME = "PBIDesktop.exe"
	for proc in psutil.process_iter():
		# check whether the process name matches
		if proc.name() == PROCNAME:
			proc.kill()
	time.sleep(3)

	# Start PBI and open the workbook
	print("Starting Power BI")
	os.system('start "" "' + WORKBOOK + '"')
	print("Waiting ",INIT_WAIT,"sec")
	time.sleep(INIT_WAIT)

	# Connect pywinauto
	print("Identifying Power BI window")
	app = Application(backend = 'uia').connect(path = PROCNAME)
	win = app.window(title_re = '.*Power BI Desktop')
	time.sleep(5)
	win.set_focus()
	
	# Refresh
	print("Refreshing")
	win.Refresh.click_input()
	win.wait("enabled", timeout = 300)
	time.sleep(5)
	print("Waiting for refresh end (timeout in ", REFRESH_TIMEOUT,"sec)")
	win.wait("enabled", timeout = REFRESH_TIMEOUT)

	# Publish
	if args.publish:
		print("Save and Publish")
		win.Publish.click_input()
		time.sleep(10)
		save_dialog = win.child_window(auto_id = "modalDialog")
		save_dialog.Save.click_input()
		time.sleep(10)
		publish_dialog = win.child_window(auto_id = "KoPublishToGroupDialog")
		publish_dialog.child_window(title = WORKSPACE, found_index=1).click_input()
		time.sleep(10)
		publish_dialog.Select.click_input()
		time.sleep(10)
		replace_dialog = win.child_window(auto_id = "KoPublishWithImpactViewDialog")
		replace_dialog.Replace.click_input()
		time.sleep(300)
		
	#Close
	print("Exiting")
	win.close()

	# Force close
	for proc in psutil.process_iter():
		if proc.name() == PROCNAME:
			proc.kill()

		
if __name__ == '__main__':
	try:
		main()
	except Exception as e:
		print(e)
		sys.exit(1)
