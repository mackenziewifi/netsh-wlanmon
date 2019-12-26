# Netsh WLAN script
A script to continuly monitor and parse the output of the 'netsh wlan show interfaces' and display the details in a windows GUI. With the option of save the details to log 

The inspiration and starting place for this script was Nigel Bowden's powershell script
# http://wifinigel.blogspot.com/2016/09/getting-data-out-of-windows-netsh-wlan.html

# Windows WLAN Data Powershell script

 Run this script from a Windows powershell (as admin) 

 Note: 
       to run Powershell scripts, you will most likely need to update
       the execution policy on your machine. Open a Powershell window
       as an adminsitrator on your machine and temporarily set the policy
       to unrestricted with this command:

       Set-ExecutionPolicy Unrestricted -scope Process

       Once your powershell window is closed, this policy change is no
       longer in effect and your machine will return to the previous
       policy. You can check your execution policy by running the 
       following command in Powershell:

       Get-ExecutionPolicy

 Note2: 
        This has only been tested on a Win 10 machine. This is a Beta script and will have bug. 

