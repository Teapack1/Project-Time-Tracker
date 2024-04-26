# Project Time Tracker

This is a straightforward tool for tracking time spent on different projects. It's designed for anyone who needs to track their project time, especially useful if your work involves handling multiple projects. The interface is easy to use: just type in your project or click project preset, the project is tracked, then click other project and the previous stops and new starts. It keeps all your monthly work neatly organized in one Excel file, with a new tab for each month.

## Features
- **Automatic Month Tab Creation**: Automatically starts a new tab in the Excel sheet each month.
- **Simple Interface**: Easy to use, especially for project-based work.
- **Monthly Excel Sheets**: Keeps all your monthly data in one file, with new tabs for each month.
- **Normalization Feature**: Adjusts days with less than 7.5 hours, filling them to 7.5.
- **Preset Project Categories**: Preset 4 cards for common projects, making it quicker to start tracking.

## Installation

### Prerequisites
- Python 3.10
- 
### Steps
1. **Download the Project**:
   - Download the ZIP file of the project from GitHub and extract it to a folder.
   - Or clone it.

2. **Install Python**:
   - Download and install Python 3.10 from the [official Python website](https://www.python.org/downloads/). Remember to check the option 'Add Python 3.10 to PATH' during installation. reboot.

3. **Install Dependencies**:
   - Open Command Prompt (CMD).
   - Navigate to the folder where you extracted the project. For example, if you extracted it to `C:\ProjectTimeTracker`, type `cd C:\ProjectTimeTracker` in CMD and press Enter.
   - Type `pip install -r requirements.txt` in CMD and press Enter to install the required libraries.

4. **Initial Setup the app**:
   - In the same CMD window, construct a command to setup your custom project tracker. See arguments bellow. Your setup command will look something like this: `python setup.py -l C:\username\Documents -p1 RnD -p2 Others -p3 Proejct3 -p4 Project4 -no 12 -w 1 -n "My Project Work Log"`.

5. **Run the Application**:
   - In the same CMD window, type `python main.py` and press Enter to start the application.


### Command-Line Arguments

- `-l`, `--excel_location`: Set the directory where the Excel file will be saved.  
  Example usage: `--excel_location "C:/Users/username/Desktop"`

- `-p1`, `--project1`: Define the name of the first quick-access project.  
  Example usage: `--project1 "Administration"`

- `-p2`, `--project2`: Define the name of the second quick-access project.  
  Example usage: `--project2 "Development"`

- `-no`, `--normalize_hours`: Set the number of hours to which each day's work should be normalized.  
  Example usage: `--normalize_hours 8`

- `-w`, `--language`: Choose the language for the application interface (0 for English, 1 for Czech).  
  Example usage: `--language 0`

- `-n`, `--name`: Provide a custom name for the Excel sheet where hours will be logged.  
  Example usage: `--name "Log"`

### Example setup:
`python setup.py -l C:\username\Documents -p1 RnD -p2 Others -p3 Proejct3 -p4 Project4 -no 12 -w 1 -n "My Project Work Log"`

## Usage
- When the application starts, enter the name of your project and click 'Activate'.
- Switch projects simply by entering a new name and clicking 'Activate' again.
- All the changes from the tracking is saved, when you hit the red cross to close the window.
- Check the Excel file for a complete record of your work, neatly organized by day and project, ready to provide for report.
- You can also use the normalize feature to fill days that does not reach 7.5 hours of projects.

