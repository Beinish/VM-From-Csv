# VM-From-Csv
This script will allow you to clone VMs from a template and give them static IP addresses, subnetmasks and default gateways.

## Spreadsheet Layout
You need to setup your .csv file exactly like the `example.csv` file I added to the repo. If you don't want to download it, the layout *has* to look like so:
Hostname | IP | Gateway | Subnet
-------- | -- | ------- | ------
Test01 | 192.168.1.1 | 192.168.1.254 | 255.255.255.0
Test02 | 192.168.1.2 | 192.168.1.254 | 255.255.255.0

Obviously you can change the headers of the columns, provided that you update them in the script as well (line 456).

## GUI Explained

![GUI](https://i.imgur.com/wgTI8H1.png)

1. VCenter - Type the name or IP of your vcenter. It will connect using your current credentials.
1. OS - Choose what OS you are planning on cloning and click "Fetch". The script will populate the rest of the fields according to your choice.
   1. Windws - The script searches for the keyword `*win*` so make sure your template has that word in it, or else it won't find it.
   1. CentOS - Same as Windows, except it searches for the keyword `*CentOS*`
   1. Red Hat - Same as the others, except it searches for the keyword `*RH*`
1. Custom Template - Make sure you use a custom template that **Doesn't** use a DHCP for it's NIC mapping. Configure one for manual IP or user input.
1. Datastore - Once you pick the datastore, the "provisioned" and "free" fields will populate with the corresponding info, showing you how much free space you have.
1. CSV - After you load the CSV, feel free to click on "calculate". It will calculate how many VMs are in that .csv file and will let you know how [much free space](https://i.imgur.com/6hWNugq.png) will be left after the cloning process.
1. CPU - Choose the amount of CPU you want.
1. RAM - Choose the amount of RAM you want.
1. VLAN - Pick the VLAN you want the VMs to be in. Make sure it corresponds with the .csv file's info (Correct gateway, IPs, etc).
1. Folder - Type a name of the folder and click "Find Folder". It will populate the dropdown menu with all the folders that match your keyword.
1. Notes - Optional of course. You can add a note to each VM you deploy.
1. Start VM after creation - Self explanatory.

When you're ready, click on "Start Clone".


### It's always a good idea to test before you use it in your environment. This script has been tested for my own needs in my own environment. Feel free to edit whatever you see fit in the script.
