# QuestionaireUI
Tkinter drop down options questionnaire with commentary from ChatGPT on the results per section and deviation from the max possible score.

it takes any Questionaire.csv file and as per the sample here, it creates an UI that has a set of questions with a set of options under the options column header in the csv file for which the weights are set.

Create your own Questionaire.csv file from the sample and modify the prompt to chat.ai API to anything that suits your use case. You can use any logo image you want to enhance the UI

# Installation
Create virtualenv and activate

virtualenv venv_tk
source venv_ds/bin/activate
# Requirements

pip install -r requirements.txt

## Program:
A GUI with dropdown options and optional help button for each question:
![Sample](https://github.com/CharlieAlphaFox/QuestionaireUI/assets/50183852/fb335314-dc17-490c-94f1-8a3ce5480054)

# Output
Once the questions are answered the program rates them with a score from about -12 to 100 according to the sample weights in the CSV file and gives a final score.

# Ai commentary

Based on the questions answers the program queries openai's API to give a report on the findings.

# Output & Possible use cases

All of the output is saved on to an .xlsx file that waits for the commentary from openai's API to finish, this can then be uploaded to a better format on Google sheets. This can be used to rate any topic on a conference/video call, get a score and produce a report.
