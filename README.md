# PARADIM ARPES Calendar
This is a script written to hep track times when instruments were in use for specific proposals. The script take a file as an input. From the input file, the script looks at all the files in the folder and determines which is the first and which is the last file. The first file is used to find starting date ajd time, and the last file is used to find the ending date and time. One event is created for the folder. The events starts at the start time and ends at the end time. The current working function is main in the ARPES_Calendar.py file. The event is created in the ARPES Calendar, which is kept by Brendan Faeth. 

##Usage 
#### Installation
First, set up a virtual environment. This script was written and tested in a conda environment. 

	conda create -n googlecalendarapi python -y
	conda activate googlecalendarapi
Then install required packages in your new environment

	pip install pandas
	pip install datetime
	pip install pytz

There is one required package that can't be easily installed with pip. That is HTMDEC Formats. HTMDEC Formats can be installed by download or cloning the repository that can be found [here](https://github.com/htmdec/htmdec_formats), and the run the command 

	pip install . 
	
from inside the repository you just downloaded or cloned. If you need test data, the folder JF24020 can be downloaded and used for test data.
### Google QuickStart Instructions
Next, follow the google quickstart instructions. Before starting the the Quickstart instructions, you must make sure one of your google accounts has been added to the application as a user. The google quickstart instructions can be found [here](https://developers.google.com/calendar/api/quickstart/python). Enabling the API and configuring the OAuth screen should be very simple and should just require clicking through a couple of screens. For the credentials step, make sure you download the file and put it in the same folder where the main.py script is and rename it to credentials.json. Finally, make sure you execute the following pip install command in your environment.

	pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib
After following these steps run the main.py script. The first time running the script should take you to a browser where you are asked to sign in to recieve your token.json file. This may take more than one attempt to execute properly. When you recieve your token.json file make sure it is in the same folder as the main.py file.

###Troubleshooting
If you are having trouble getting things working try running the testing.py file. This file is apart of the google quickstart instructions and should be easier to get working. If you have that working and are still having trouble with the ARPES_Calendar file, the issue is likely with HTMDEC Formats or possibly even your data. 