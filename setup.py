import argparse
import json

def parse_arguments():
    parser = argparse.ArgumentParser(description="Setup for Project Time Tracker")
    parser.add_argument("-l", "--excel_location", type=str, default=".", help="Directory to save the Excel file ex.: C:/Users/username/Desktop")
    parser.add_argument("-p1", "--project1", type=str, default="Others", help="Name of the first quick project")
    parser.add_argument("-p2", "--project2", type=str, default="RnD", help="Name of the second quick project")
    parser.add_argument("-no", "--normalize_hours", type=float, default=7.5, help="Number of hours to normalize to every day")
    parser.add_argument("-w", "--language", type=int, default=0, choices=[0, 1], help="Language selection (0 - English, 1 - Czech)")
    parser.add_argument("-n", "--name", type=str, default="projectLog", help="Name the sheet file 'Hours Loggging Sheet Name'")
    return parser.parse_args()

def save_configuration(args):
    config = {
        "excel_location": args.excel_location,
        "default_projects": [args.project1, args.project2],
        "normalize_hours": args.normalize_hours,
        "language": args.language,
        "name": args.name
    }
    with open('config.json', 'w') as config_file:
        json.dump(config, config_file)

def main():
    args = parse_arguments()
    save_configuration(args)
    
    
    
    print("\n")
    print(f"Excel file named '{args.name}' set to location: '{args.excel_location}'")
    print(f"Default projects: '{args.project1}', '{args.project2}'")
    print(f"Hours will be normalized to: '{args.normalize_hours}' hours")
    print(f"Language: '{args.language}' (0 - English, 1 - Czech))")
    print("\n")
    print("Setup completed. Configuration saved.")
    print("To run the program, run main.py")

if __name__ == "__main__":
    main()
