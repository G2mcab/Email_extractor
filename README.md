To clone the program use the following command

git clone https://github.com/G2mcab/Email_extractor.git

Or You can download the program here.

Once the process is finished

You can

1. Open launch.bat
2. Use command line

Launch.bat 
If you are on windows machine you can use the launch.bat to launch the different program modules it will also automatically install libraries

Using command Line

Open CMD
then type 

python -m venv env
env\Scripts\Activate
pip install -r Requirements.txt

To extract simple csv

python -m Simple_extractor

To use the advanced CLI extractor

python -m Full_extractor

To launch the GUI version

python -m Full_extractor_GUI

To launch the advanced Email extractor
python -m Advanced_Email_extractor

Features

Simple extractor
- Enter the sender email account
- Exports the emails into emails_from_{sender_email}.csv
- Also have option to export and delete
- Export and archieve emails

The file will be in the same folder as the program

Advanced CLI extrator
- does the same thing as above with more error handling capabilites, thread, logging capabilities

GUI version
- Added GUI capabilities

Advanced Version

- Extracts into Emails folder
- then creates a subfolder with the format emails_from_{sender_email}
- In the folder there are three files HTML, csv and json
- If you open the JSON file you can find the emails from that specific sender formatted in calender
