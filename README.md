# Companion of Darkness Automatic Grammar Patcher

This program automatically applies proofreading fixes from the google sheets proofreading spreadsheet to the game's Ren'Py script (`script.rpy`).
Everything is handled in the single `main.py` file. And the code is quite unpleasant to look at (read: HORRIBLE). But hey, it works... (Or atleast I think I haven't checked every line of the output but it looks ok?)

## Features

- Automatically creates an enviroment on runtime.
- Creates a new `modscript.rpy` file containing the updated script.
- WIP: Diff mode, get how many lines are changed without changing them.
- WIP: Advanced modes: Manually configure parameters for patching.
- 
## Requirements

- Python 3.6+
- Ren'Py script (`script.rpy`)
- The google sheets spreadsheet downloaded in .xlsx format.

## Usage

1. Place this script (`main.py`) in a new directory.
2. Copy your `script.rpy` file into the same directory.
3. Go to google sheets and download a copy of the proofreading spreadsheet in the .xlsx (Microsoft Excel) format.
4. Place this spreadsheet in the same directory as the program.
5. Run the script:
   ```bash
   python main.py
   ```
6. Follow the instructions given in the command line. (Currently only quick patching is implemented)
7. The modified Ren'Py script will be saved as `modscript.rpy`.

## Notes

- If more than 25 lines are not found, the script will exit with a critical warning. Ensure the spreadsheet is accurate.
- In case of issues or excessive amounts of missed lines, manually review the changes or open a GitHub issue or complain to the spreadsheet maintainers.
