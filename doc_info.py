head = '''###host_name### NAC Implementation\n DATE   START_TIME till END_TIME\nChange #'''
first_line = "Participants & Contact Information:"

second_line = 'Objectives'
third_line = 'Procedure'
numbered_line_part_one = ['Conduct pre job brief',
"Contact the NOC (network alarms)", 
"Contact the Help Desk           ",
"Apply spanning tree and Access List changes on ###host_name###" 
]
numbered_line_part_two = ['Prior to configuration, add ###host_name### to ISE',
'Begin configuration changes to ###host_name### closet switch']

numbered_line_part_three = [ 'Check NAC/ISE console: remediate any items that show up in monitor mode profile',

 ]                                                                     


span_info = """Spanning-tree / VTY & SNMP ACLs  global configuration:

NOTE:
Replace the ISE Device Group community string for ###host_name### from Atm05Pher1c  to  Mnt5tHe1enS"""
