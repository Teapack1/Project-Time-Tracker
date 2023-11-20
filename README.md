# Project Time Tracker

## Introduction
This is a straightforward tool for tracking time spent on different projects. It's designed for anyone who needs to manage their project time, especially useful if your work involves handling multiple projects. The interface is easy to use: just type in your project, click a button, and the tracker starts. It keeps all your monthly work neatly organized in one Excel file, with a new tab for each month.

## Features
- **Automatic Month Tab Creation**: Automatically starts a new tab in the Excel sheet each month.
- **Simple Interface**: Easy to use, especially for project-based work.
- **Monthly Excel Sheets**: Keeps all your monthly data in one file, with new tabs for each month.
- **Normalization Feature**: Adjusts days with less than 7.5 hours of work to total 7.5 hours, balancing your work schedule.
- **Predefined Project Categories**: Two default categories for common projects, making it quicker to start tracking.

## Installation for Non-Coders

### Prerequisites
- Python 3.10
- Python libraries (included see Install Dependencies)

### Steps
1. **Download the Project**:
   - Download the ZIP file of the project from GitHub and extract it to a folder.

2. **Install Python**:
   - Download and install Python 3.10 from the [official Python website](https://www.python.org/downloads/). Remember to check the option 'Add Python 3.10 to PATH' during installation.

3. **Install Dependencies**:
   - Open Command Prompt (CMD).
   - Navigate to the folder where you extracted the project. For example, if you extracted it to `C:\ProjectTimeTracker`, type `cd C:\ProjectTimeTracker` in CMD and press Enter.
   - Type `pip install -r requirements.txt` in CMD and press Enter to install the required libraries.

4. **Run the Application**:
   - In the same CMD window, type `python main.py` and press Enter to start the application.

## Usage
- When the application starts, enter the name of your project and click 'Activate'.
- Switch projects simply by entering a new name and clicking 'Activate' again.
- All the changes from the tracking is saved, when you hit the red cross to close the.
- At the end of the month, check the Excel file for a complete record of your work, neatly organized by day and project.
- You can also use the normalize feature to fill days that does not reach 7.5 hours of projects.
- You can adjust the look of your sheets by editing the "Project_Sheet.xlsx" template in the folder.
- Users can run this script from the command line to set up their preferences. For example:
    `python setup.py --excel_location "C:/MyTimeTracker" --project1 "R&D" --project2 "Admin" --normalize_hours 8`


## License
This project is released under the [MIT License](LICENSE).
