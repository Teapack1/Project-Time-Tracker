import csv

class Data:

    def __init__(self):
        self.database = []

    def write_data(self, proj, tim):
        with open("project_data.txt", mode="a", newline="") as file:
            writer = csv.writer(file)
            writer.writerow([proj, tim])
            print(f"Written: {proj}; {tim}")

    def read_data(self):
        with open("project_data.txt", mode="r", newline="") as file:
            reader = csv.DictReader(file)
            for row in reader:
                dict = {}
                dict["project"] = row["project"]
                dict["time"] = row["time"]
                self.database.append(row)
        return self.database

    def remove_last(self):
        with open("project_data.txt", mode="r+") as file:
            lines = file.readlines()[-1]
            print(f"last written {lines}")

    def delete_data(self):
        with open("project_data.txt", "r") as infile:
            # Read the first row of the file
            first_row = infile.readline()

        with open("project_data.txt", "w") as outfile:
            # Write the first row back to the file
            outfile.write(first_row)