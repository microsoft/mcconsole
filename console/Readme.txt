McConsole 

Goal: Help users remember and run commands, and assit debug.

-----------------
Instructions:
0. Data Collection: The McConsole collects NO data from user sut folder, NO logs is saved to server either. Mobaxterm Free Version is leveraged in Storing it's password. McConsole does not interact with Mobaxterm to exchange password. 
1. Change/Check sut.{sut_ip}.setting correctly reflect your various consoles IP and login info. 
2. Anything in above file with {CONSOLE_TYPE}_IP: will be considered a console and be opened in windows terminal as a tab. 
3. Edit {SKU_TYPE}_{CONSOLE_TYPE}_COMMANDS.csv for commands for each console type that you want to remember. 
4. Double click helper window to run a a cmd, or type in the console as normal
5. Right click helper window to modify a cmd.
6. Logs are saved in console folder. 
10. You can put the sound files under 'sound' folder when an command is clicked, also change BGM under the 'BGM' subfolder. If you run this on jumpbox, RDP can pass the sound output to labtop too. Enjoy!


-----------------
Change log:
v1.0.2: Initial public release


