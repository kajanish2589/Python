import os,time
from py3270 import Emulator
from function import *


user_id = "CNOALN"
pwd_temp = "123LOVE"
session_id = 'TLIFE'


#==================================================================================================================================
# Get the path of the currently executed Python file
current_file_path = os.path.abspath(__file__)
# Create a new html file in the same directory
html_save_screen_file_path_out = os.path.join(os.path.dirname(current_file_path), 'Policy_add_Screenshot.html')
with open(html_save_screen_file_path_out, 'w') as new_file:
    new_file.write(' ')
# Correctly construct the path to Screenshot.html
htmlf = '"'+os.path.join(os.path.dirname(current_file_path), 'Policy_add_Screenshot.html')+'"'
# Correctly construct the path to DATA.xlsx
excelf = os.path.join(os.path.dirname(current_file_path), "DATA.xlsx")        
#excel_data_load(excelf)
record = excel_data_load(excelf) 
#==================================================================================================================================

#===============================================================================================
em = Emulator(visible=False, args=["-trace","-tracefile","run.log"]) #call emulator method
em.connect('mf.conseco.com:23') #connect to host

em.exec_command(b"Wait(Output)")

em.save_screen(htmlf) 
string_wait(em,"===> ")
em.save_screen(htmlf) 
em.fill_field(24,7,'TPX',3)
em.save_screen(htmlf) 
em.send_enter()
em.save_screen(htmlf) 
em.fill_field(14,20,user_id.strip(),7)
em.fill_field(15,20,pwd_temp.strip(),7)
em.save_screen(htmlf) 
em.send_enter()
em.save_screen(htmlf) 
print("login successfully")
string_wait(em,"Command ===> ")
em.save_screen(htmlf) 
em.send_string(" " + session_id + "\\n")
em.save_screen(htmlf)       
em.send_enter()
print("Enter into " + session_id )
em.save_screen(htmlf) 
string_wait(em,"????")
em.send_string("VTGU")
em.save_screen(htmlf)  
em.send_enter()
em.save_screen(htmlf) 
em.wait_for_field()
em.save_screen(htmlf)
print(record)

if not record:
    print("The Excel Data empty. Stored Location ==>" + excelf)
else:
    for ea_data in record:   
        policy_add_screen_fill(ea_data,em,htmlf)
        error_check(ea_data,em,htmlf) 
        ea1_screen_fill(ea_data,em,htmlf)
        ea2_screen_fill(ea_data,em,htmlf)
        ea3_screen_fill(ea_data,em,htmlf)
        ea4_screen_fill(ea_data,em,htmlf)
        ea5_screen_fill(ea_data,em,htmlf)
        ea6_screen_fill(ea_data,em,htmlf)
        ea7_screen_fill(ea_data,em,htmlf)
        ea8_screen_fill(ea_data,em,htmlf)
        ea9_screen_fill(ea_data,em,htmlf)
        ea10_screen_fill(ea_data,em,htmlf)
        em.save_screen(htmlf) 
        em.send_enter()
        em.save_screen(htmlf) 
        ea11_screen_fill(ea_data,em,htmlf)
        em.save_screen(htmlf) 
        em.send_pf5()
        em.save_screen(htmlf) 
        string_wait(em,"DATABASE UPDATED")
        em.save_screen(htmlf) 
        print(em.string_get(22,48,30))
        em.send_pf2()
        time.sleep(5)
        rider_add_screen_fill(ea_data,em,htmlf)
        GS_screen_fill(ea_data,em,htmlf)  
        policy_complete_add_screen_fill(ea_data,em,htmlf)
        time.sleep(5)
        string_wait(em,"POLICY COMPLETE")
        print(em.string_get(19,48,30))
        em.send_pf2()
        time.sleep(5)



em.save_screen(htmlf) 
#==================================================================================================================================
em.terminate()  #disconnect to host


