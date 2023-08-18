# version: 1.2
import asyncio
import traceback
import tkinter as tk
from time import sleep
from tkinter import ttk, Text
from tkinter import messagebox
from functools import partial
from PIL import Image, ImageTk
import pandas as pd
from retrying import retry
import xlsxwriter
import webbrowser
import requests
import json
import re

# Load the data from the CSV file
data = pd.read_csv('Questionaire.csv')

# Initialize an empty dictionary to hold the questions and options
questions = {}

@retry(stop_max_attempt_number=2, wait_fixed=5000)
async def chatbot_answer(startup_name, responses, score):
    print('Sending request to OpenAI...')
    try:
        api_endpoint = "https://api.openai.com/v1/chat/completions"
        headers = {"Authorization": "Bearer sk-yourAPIkeysGoHereLikeThisOrWithImportedOs", 'Content-Type': 'application/json'}

        # Prepare a list of prompts for the AI.
        prompts = [f"{key}: {value['Response']}, Top Score Difference: {value['Score Difference']}" for key, value in responses.items() if value['Response']]
        prompts = '\n'.join(prompts)

        txt1 = "You are an expert on Entrepreneurship with an impartial view and very opinionated on the recommendations about startups." # persona
        txt2 = f"The responses to section: questions and difference from top score (the lower the better) about {startup_name} were as follows:\n{prompts}\n with a final score of {score} of 100\n"
        txt3 = "Please note the final score thresholdsare : <32: Not a good idea, 32-40: Might work but needs attention, 40-60: A good idea that needs work, 61-76 is great, >76: Excellent idea. To expand"
        txt4 = "Using your expertise, please generate a very comprehensive, well formated and lenghty report article per question, excluding the Info section, and based on the score difference: describing any threats, the startup potentials on how each area can be improved, scaled, delivered, optimized, " 
        txt5 = "the risks and the concrete next steps, the possible product phase innovation ideas and ramifications, as well as an overall conclusion, all based on your expertise and the score difference"
        
        ttxt = txt2 + txt3 + txt4 + txt5 # Excluding txt1 as its the persona
        data = {
            "model": "gpt-4-0613",
            "max_tokens": 2950,
            "temperature": 0.4,
            # "top_p": 1,
            "frequency_penalty": 1,
            "messages": [
                {
                    "role": "system",
                    "content": txt1
                },
                {
                    "role": "user",
                    "content": ttxt
                }
            ],
        }
        # https://beta.openai.com/docs/api-reference/completions/create
        # https://atoz.ai/articles/understanding-temperature-top-p-presence-penalty-and-frequency-penalty-in-language-models-like-gpt-3/

        response = requests.post(api_endpoint, headers=headers, data=json.dumps(data))
        print('Response received, saving to file...')

        response.raise_for_status()
    except Exception as e:
        print(traceback.format_exc())
        print(e)
        response_json = response.json()
        print(response_json.get('choices', [{}])[0].get('content', '').strip())
        return None
    
    try:
        response_json = response.json()
        response_content = response_json['choices'][0]['message']['content'].strip()
        print(response_content)
        return response_content
    except Exception:
        print("Error decoding JSON from response.")
        print("Response:", response.text)
        return None

# Iterate over each data_row in the dataframe
for index, data_row in data.iterrows():
    if pd.isna(data_row['Task']):
        continue

    # The 'Section' and 'Task' fields form the key in the dictionary, separated by a colon
    key = f"{data_row['Section']}: {data_row['Task']}"
    
    # The 'Options' field is split into a list on the , character (",")
    if pd.isna(data_row['Options']):
        options = []
        weights = []
    else:
        options = [opt.strip() for opt in data_row['Options'].split(', ')]
        weights = [int(wt.strip()) for wt in data_row['Weights'].split(',')]
    
    # Add this question to the dictionary
    questions[key] = {
    'question': data_row['Task'],
    'options': options,
    'help_link': data_row['Help_Link'] if 'Help_Link' in data_row else None,
    'weights': weights
}

# Print out the resulting dictionary
for key, value in questions.items():
    print(f"Section/Task: {key}")
    print(f"Question: {value['question']}")
    print(f"Options: {value['options']}\n")

responses = {}

def proper_filename(filename):
    return re.sub(r'[\\/*?:"<>|]', "_", filename)

async def submit():
    global startup_name, score
    def write_to_xls(data_dict, xls_path, text):
        global df
        # Convert dictionary to a DataFrame
        df = pd.DataFrame.from_dict(data_dict, orient='index')
        # Reset the index
        df.reset_index(inplace=True)
        try:
            name = str(xls_path)
            f_path = str(proper_filename(xls_path) + '.xlsx')
            writer = pd.ExcelWriter(f_path, engine="xlsxwriter")
            # Define the format properties for the header
            workbook  = writer.book
            header_format = workbook.add_format({
                'bold': True,
                'text_wrap': True,
                # 'valign': 'top',
                'fg_color': '#D7E4BC',
                'border': 1
            })
            # Write DataFrame to excel
            df.to_excel(writer, index=False, sheet_name=name) 
            # Get the worksheet object
            worksheet = writer.sheets[name]
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)
            # Write additional text
            worksheet.write(len(df) + 1, 0, text)
            writer.close()
        except IOError as e:
            print("I/O error:")
            print(e)
    
    score = 0
    for key, value in questions.items():
        question_score = 0
        response = value['var'].get()
        if value['weights']:  # This will be False if the list is empty
            max_possible_score = max(value['weights'])  
        else:
            max_possible_score = 0  # Default value if no weights are provided
        if response in value['options']:
            question_score = value['weights'][value['options'].index(response)]
        elif response == 'Other (specify)':
            question_score = 1  # or any other default score you want to assign
        score_difference = max_possible_score - question_score  # Calculate the difference between the max possible score and the actual score
        score += question_score
        responses[key] = {
            'Response': response,
            'Score': question_score,
            'Score Difference': score_difference,  # Add the score difference to the responses dictionary
        }
    scoredif = 100 - score
    responses['Total Score'] = {'Response': '', 'Score': score, 'Score Difference': scoredif}  # No score difference for total score

    messagebox.showinfo(f"Responses saved, preparing a report\nYour score is: {score}")
    print(f"Responses saved, preparing a report!\nYour score is: {score}")
    
    with open('responses.json', 'w') as f:
        json.dump(responses, f, indent=4)
    
    startup_name = 'responses'  # default name
    keywords = ['name', 'Venture', 'Startup', 'Concept']
    for key, response in responses.items():
        if any(keyword in key for keyword in keywords):
            if response['Response']:  # Check that response is not an empty string
                startup_name = response['Response']
                break
    text = await chatbot_answer(startup_name, responses, score)
    write_to_xls(responses, startup_name, text)
    with open(f'{startup_name}.json', 'w') as f:
        json.dump(responses, f, indent=4)

# Add help button
def open_help_link(link):
    webbrowser.open(link)

root = tk.Tk()
root.title("Ridddec MVP - Venture Evaluation")

# Set a dark theme for the GUI
font_size = 18
root.configure(bg='SandyBrown')
style = ttk.Style(root)
style.theme_use("clam")
style.configure('.', background='white', foreground='grey')
style.configure('TLabel', foreground='grey', font=('Helvetica', font_size))
style.configure('TButton', foreground='white', background='light grey')
style.configure('TCombobox', fieldbackground='white', background='white', foreground='white', font=('Helvetica', font_size))
style.configure('TEntry', foreground='black')
style.configure('help.TButton', padding=(2, 2, 2, 2))  # adjust these numbers as needed
root.geometry("1400x800")  # You can set this to your preferred size

text = Text(root, height=22, font=('Helvetica'))

# Load the logo (the file logo.png should be in the same directory as your script)
logo_image = Image.open("logo.png")
logo_photo = ImageTk.PhotoImage(logo_image)
logo_label = ttk.Label(root, image=logo_photo)
logo_label.image = logo_photo  # keep a reference to the image object to prevent it from being garbage collected
logo_label.grid(row=1, column=0, columnspan=1)  # set the logo at the top of the window

def next_section(current_frame, next_frame):
    # Hide current frame and show the next frame
    current_frame.grid_remove()
    next_frame.grid()

# Build a list of unique sections in the correct order
sections = []
current_section = None
for key, value in questions.items():
    section = key.split(":")[0]
    if section != current_section:
        sections.append(section)
        current_section = section

# Create a frame for each section
frames = {section: ttk.Frame(root, padding="20 60 20 60") for section in sections}

# Initially hide all frames except the first one
count = 0
for frame in frames.values():
    frame.grid(row=3, column=0, sticky=(tk.N, tk.S)) # , tk.W, tk.E 
    frame.grid_remove()
    count += 1
frames[sections[0]].grid()

# Now, fill in the frames with questions
current_section = None
row = 2
for key, value in questions.items():
    section = key.split(":")[0]
    if section != current_section:
        # Add Next button to the previous section
        if current_section is not None:
            row += 2
            next_section_key = sections[sections.index(current_section) + 1] if sections.index(current_section) + 1 < len(sections) else None
            if next_section_key is not None:
                next_command = partial(next_section, frames[current_section], frames[next_section_key])
                ttk.Button(frames[current_section], text="Next", command=next_command).grid(column=1, row=row, sticky=tk.E)

                # Add Back button to the previous section
                back_section_key = sections[sections.index(current_section) - 1] if sections.index(current_section) > 0 else None
                if back_section_key is not None:
                    back_command = partial(next_section, frames[current_section], frames[back_section_key])  # Yes, it's still next_section, because next_section just hides the current frame and shows the specified frame
                    ttk.Button(frames[current_section], text="Back", command=back_command).grid(column=0, row=row, sticky=tk.W)

        current_section = section
        # Add section heading
        ttk.Label(frames[section], text=section + ' Section', font=('Helvetica', 20)).grid(column=0, row=0, columnspan=2, sticky=tk.W)
        row = 1
    
    ttk.Label(frames[section], text=value['question']).grid(column=0, row=row, sticky=tk.W)
    value['var'] = tk.StringVar()
    if value['options']:
        # Add "Other (specify)" option to the list of options
        value['options'].append('Other (specify)')
        widget = ttk.Combobox(frames[section], textvariable=value['var'], values=value['options'], width=50)
        widget.grid(column=1, row=row, sticky=(tk.W, tk.E))

        # If "Other (specify)" is selected, display an Entry box for custom input
        def on_combobox_changed(event, var=value['var'], frame=frames[section], row=row):
            if var.get() == 'Other (specify)':
                var.set('')  # Clear the variable
                entry = ttk.Entry(frame)
                entry.grid(column=2, row=row, sticky=(tk.W, tk.E))

        widget.bind('<<ComboboxSelected>>', on_combobox_changed)
    else:
        if section == sections[0] and row == 1:  # this is the first question of the first section
            widget = tk.Text(frames[section], height=8, width=60)
            widget.grid(column=1, row=row, sticky=(tk.W, tk.E))
        else:
            widget = ttk.Entry(frames[section], textvariable=value['var'])
            widget.grid(column=1, row=row, sticky=(tk.W, tk.E))
    
    # Add help button if there is a help link, after question is placed <-- CHANGE HERE
    if isinstance(value['help_link'], str) and value['help_link'].strip() != '':
        help_button = ttk.Button(frames[section], text="?", command=partial(open_help_link, value['help_link']), style='help.TButton')
        help_button.grid(column=2, row=row, sticky=tk.W)
    
    row += 1

def close_window():
    root.destroy()

# Add Exit button to the last frame and remove the 'Next' button
if current_section is not None:
    row += 4
    # Add Back button to the previous section
    back_section_key = sections[sections.index(current_section) - 1] if sections.index(current_section) > 0 else None
    if back_section_key is not None:
        back_command = partial(next_section, frames[current_section], frames[back_section_key])  # Yes, it's still next_section, next_section hides the current frame and shows the needed frame
        ttk.Button(frames[current_section], text="Back", command=back_command).grid(column=0, row=row, sticky=tk.W)
    ttk.Button(frames[current_section], text="Submit", command=lambda: [asyncio.run(submit()), close_window()]).grid(column=1, row=row, sticky=tk.E)

root.mainloop()
