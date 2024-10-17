import os
import subprocess
import sys
import platform
import re
import time

try:
    from colorama import Fore, Back, Style, init
    import pandas as pd
except ImportError:
    print(
        "Found uninstalled imports, this should fix itself after the automatic restart. If not, open an issue on github."
    )

def create_venv_if_needed():
    venv_dir = os.path.join(os.getcwd(), "env")

    if not os.path.exists(venv_dir):
        print("Virtual environment not found. Creating...")

        subprocess.check_call([sys.executable, "-m", "venv", "env"])

        print("Installing dependencies...")

        if platform.system() == "Windows":
            pip_path = os.path.join(venv_dir, "Scripts", "pip.exe")
        else:
            pip_path = os.path.join(venv_dir, "bin", "pip")

        subprocess.check_call([pip_path, "install", "pandas", "openpyxl", "colorama"])
        print("Dependencies installed!")


def run_program_in_venv():
    venv_dir = os.path.join(os.getcwd(), "env")

    if platform.system() == "Windows":
        python_executable = os.path.join(venv_dir, "Scripts", "python.exe")
    else:
        python_executable = os.path.join(venv_dir, "bin", "python")

    subprocess.check_call([python_executable, *sys.argv])


def main():
    unereplaced_lines = 0
    create_venv_if_needed()
    if sys.prefix != sys.base_prefix:
        init(autoreset=True)
        print("Initialisation completed. Running rest of program.")

        patch_mode = (
            input(
                "Please select the patch mode"
                " ("
                "q = quick mode, d = diff mode, a = advanced mode, r = raw diff mode, default = q"
                "): "
            )
            .lower()
            .strip()
        )

        if patch_mode not in ["q", "d", "a", "r"]:
            print("Option not found, defaulting to quick patch mode.")
            patch_mode = "q"
        elif patch_mode in ["a", "r", "d"]:
            print(
                "NOTICE: DIFF MODE, ADVANCED PATCH MODE AND RAW DIFF MODE ARE CURRENTLY NOY IMPLEMENTED. "
                " THE PROGRAM WILL NOW EXIT."
            )
            sys.exit(0)
        xlsx_file = input(
            "Please enter the relative path to the xlsx spreadsheet (default: input.xlsx): "
        )

        if xlsx_file == "":
            xlsx_file = "input.xlsx"

        output_folder = "sheets_csv"
        os.makedirs(output_folder, exist_ok=True)
        os.makedirs("output", exist_ok=True)

        xls = pd.ExcelFile(xlsx_file)

        for sheet_name in xls.sheet_names:
            if sheet_name not in ["Overview", "Characters"]:
                df = pd.read_excel(xlsx_file, sheet_name=sheet_name)

                csv_file = os.path.join(output_folder, f"{sheet_name}.csv")
                df.to_csv(csv_file, index=False)

                print(f"Saved {sheet_name} to {csv_file}")

        print("Processing CSV files to output directory...")
        csv_dir = os.path.join(".", "sheets_csv")
        all_sheets = []

        for sheet_name in sorted(os.listdir(csv_dir)):
            active_sheet = pd.read_csv(os.path.join(csv_dir, sheet_name), header=None)
            active_sheet = active_sheet[active_sheet[0] != "True"]
            active_sheet = active_sheet[
                active_sheet[3].notna() & (active_sheet[3] != "")
            ]
            active_sheet = active_sheet.drop([0, 2, 7, 8], axis=1)
            active_sheet = active_sheet.iloc[1:]
            all_sheets.append(active_sheet)

        final_sheet = pd.concat(all_sheets, ignore_index=True)
        final_sheet.to_csv(
            os.path.join("output", "full_script.csv"), index=False, header=False
        )
        print("Patching...")
        start_time = time.time()

        def replace_lines(csv_file, txt_file):
            patch_blocked = False
            lines_protected = 0
            unereplaced_lines = 0
            total_lines = 0
            lines_found = 0

            df = pd.read_csv(csv_file)

            search_strings = df.iloc[:, 1].tolist()

            def clean_string(s):
                return re.sub(r"[^a-zA-Z0-9]", "", s).lower()

            def replace_line(s, replacement):
                pattern = r'"(.*)"'
                modified_string = re.sub(pattern, f'"{replacement}"', s)
                return modified_string

            with open(txt_file, "r") as f:
                lines = f.readlines()

            # This is remmnant of the original idea to perform a bidirectional search in order to save time.
            # Didn't work :P (Or atleast I'm too lazy to do it right)
            start_index = 0

            for search_string in search_strings:
                found = False
                total_lines += 1

                cleaned_search_string = clean_string(search_string)

                for i in range(start_index, len(lines)):
                    cleaned_line = clean_string(lines[i])

                    # if (
                    #   search_string in lines[i]
                    #   or cleaned_search_string.strip().lower() in cleaned_line.strip().lower()
                    # ):

                    if (
                        cleaned_search_string.strip().lower()
                        in cleaned_line.strip().lower()
                    ):
                        # print(search_string in lines[i])
                        # print(cleaned_search_string in cleaned_line)
                        # print(f'String "{search_string}" found on line {i + 1}')
                        if i + 1 < 960:
                            print(f"Prevented patcher from patching text into protected area on line {i+1}.")
                            lines_protected += 1
                            if patch_blocked == False:
                                lines_found += 1
                            patch_blocked = True
                            start_index = i + 1
                            continue

                        if type(df.iloc[total_lines - 1, 4]) == type("p"):
                            replaced_string = df.iloc[total_lines - 1, 4].split(
                                "\n", 1
                            )[0].replace('"', '\\"')
                            lines[i] = replace_line(lines[i], replaced_string)
                            start_index = 0
                            patch_blocked = False
                        elif type(df.iloc[total_lines - 1, 3]) == type("p"):
                            replaced_string = df.iloc[total_lines - 1, 3].split(
                                "\n", 1
                            )[0].replace('"', '\\"')
                            lines[i] = replace_line(lines[i], replaced_string)
                            start_index = 0
                            patch_blocked = False
                        elif type(df.iloc[total_lines - 1, 2]) == type("p"):
                            replaced_string = df.iloc[total_lines - 1, 2].split(
                                "\n", 1
                            )[0].replace('"', '\\"')
                            lines[i] = replace_line(lines[i], replaced_string)
                            start_index = 0
                            patch_blocked = False
                        else:
                            print(
                                f'No replacement for "{search_string}" found on line {i + 1}'
                            )
                            start_index = 0
                            patch_blocked = False
                            unereplaced_lines += 1

                        found = True
                        lines_found += 1

                        break

                if not found:
                    print(Back.YELLOW + Fore.RED + Style.BRIGHT + 
                        f'String "{search_string}" was not found in the script.rpy file'
                    )
            print('\nModified script exported to patched_script.rpy\n')
            print(f"\nTotal lines processed: {total_lines}")
            print(f"\nLines found: {lines_found}")
            print(f"Number of lines not found: {total_lines - lines_found}")
            print(f"\nNumber of lines protected: {lines_protected}")

            with open("patched_script.rpy", "w") as file:
                for line in lines:
                    file.write(line)

            if total_lines - lines_found >= 20:
                print(
                    "\nCRITICAL: MORE THAN 20 LINES WEREN'T FOUND!\n"
                    "THAT MEANS SOMETHING PROBABLY WENT WRONG!\n"
                    "DELETE ALL GENERATED FILES AND RESTART THE PROGRAM.\n"
                    "IF THE ISSUE PERSISTS, OPEN AN ISSUE ON GITHUB.\n"
                    "THE PROGRAM WILL NOW EXIT."
                )
                sys.exit(1)

            elif total_lines != 0:
                print(
                    "\nWarning: Non-zero number of lines were not found in the script.\n"
                    "This could be purely because of mistakes made by the spreadsheet maintainers.\n"
                    "You should manually go over every missed line and make sure that there are no errors."
                )
            else:
                print("Congrats! You somehow managed to patch more lines than processed... That's... Good? I think? You should be file?")

            print(f'\n{unereplaced_lines} Lines were not replaced due to missing replacements. This is probably not a problem.\n')
        csv_file = os.path.join("output", "full_script.csv")
        txt_file = "script.rpy"
        replace_lines(csv_file, txt_file)

        print(
            "\nPatch complete (It should be complete, if any errors are found, open an issue on GitHub)!"
        )
        print("\n--- %s seconds ---" % (time.time() - start_time))
        print("or")
        print("--- %s minutes ---" % ((time.time() - start_time) / 60))

    else:
        print(
            "Not in virtual environment. Restarting inside the virtual environment..."
        )
        run_program_in_venv()


if __name__ == "__main__":
    main()
