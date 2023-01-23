# MS-Office
Contains instruction for installation

A. MS Office installation Guide:
1. Download the Office Deployment Tool from the Microsoft Download Center\
Link: https://www.microsoft.com/en-us/download/details.aspx?id=49117 <br />
(Run the tool and extract it to a location on your system)
2. Create the configuration.xml file\
Link: https://config.office.com/deploymentsettings <br />
(Place the configuration file in the same location as in step 1)
3. Office Installation:
(Run the command prompt as admin and navigate to location as in step 1)\
Run the command: setup /configure configuration.xml\
The installation should start.

B. Video Guide to install MS Office Professional Plus LTSC 2021: https://youtu.be/5J7U-imkcyM

C. Activation
1. GVLK for office 2016/2019/2021 (xxxxx-xxxxx-xxxxx-xxxxx-xxxxx)<br />
Link: https://learn.microsoft.com/en-us/DeployOffice/vlactivation/gvlks?redirectedfrom=MSDN<br />
2. Offline KMS host address list:<br />
  (Check KMS_host.txt or search online)<br />
3. Run command prompt as admin and navigate to one ofthe location where OSPP.vbs file is present (check in your system)\
Path 1: System_drive\Program Files\Microsoft Office\Office16\
Path 2: System_drive\Program Files (x86)\Microsoft Office\Office16<br />
4. Run the following command one by one\
   cscript ospp.vbs /inpkey:[Product-GVLK-from-step-1]\
   cscript ospp.vbs /sethst:[kms-host-address-from-step-2]\
   cscript ospp.vbs /act\
   cscript ospp.vbs /dstatus\
(Note: If some host address do not work, then try other host address in some time)

