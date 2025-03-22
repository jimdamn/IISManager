# IISManager
PowerShell script that implements a GUI for basic IIS management on multiple networked servers. It was designed for support personnel who may be less comfortable with command line management.  

The script is very specific to my teams duties, but can be freely modified to your preference.

The script allows you to perform IIS Resets on multiple hosts without having to login to each host.  
It also allows you to query the host for application pools in IIS, and then perform application pool recycles on multiple hosts.
The script allows you to set scheduled application pool recycles on multiple hosts.
Lastly, it creates and updates log files for each of the tasks performed using the script, including which action and on what hosts it was peformed.

This version of the script assumes a number of things:  
-You have enabled remote scripting on the target hosts.
-You maintain production and test environments, which are used in the naming convention of the hosts. (the script may need to be modified to suit your needs)
-You use a common naming convention for host names. (the script may need to be modified to suit your needs)

