global_config = """

 This is where you put your globacl config
 !
  """

per_port_part1 = """
 !This is where you put the 1st part of the port config, the program spits out the data vlan after this
 shut
 authentication event server dead action reinitialize vlan """
 
per_port_part2 = """
 
 !The 2ed part of the per-port config
 no shutdown
 !\n
 """

helper_addresses_end = """
  !IP helpers for under the data vlan
 ip helper-address 10.0.0.0
!
 """

snooping_config_end = """
 no ip dhcp snooping information option
 ip dhcp snooping
 !
 """

