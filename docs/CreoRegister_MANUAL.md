# CreoRegister - User Manual (v1.0)

CreoRegister writes a new Creo part or assembly into the shared
**Part Number Register** with one click. You don't need to scroll through the
register or type the description twice.

It handles **M** (Manufactured parts), **A** (Assemblies) and **B** (Bought
parts). Elec, Fasteners and Documents are still filled in by hand.

## How the part numbers work (read this once)
- **M and A:** letter + 2-digit project ID + 4-digit number, for example
  `M281975` = Manufactured part, project 28, number 1975.
  The 4-digit number is **shared by all projects** on a sheet, so each number
  exists once only. Always take the **next free** number shown by the tool.
- **B:** the Bought part number as listed in the register, for example `B002678`.
  There is no project ID.
- The project ID must exist on the register's **Index** sheet.

## Setup (once per PC, about 5 minutes)
1. Put `CreoRegister.xlsm` on your PC. Use your own copy, not one shared copy:
   the tool saves its log into its own file.
2. Open it and click **Enable Content** if Excel asks.
3. Alt+F11 > Tools > References > tick **Creo VB API Type Library**
   (exactly ONE Creo library). Close the editor.
   - Your Creo VB API must already work on this PC: the API is installed and
     registered, and the `PRO_COMM_MSG_EXE` variable is set. Ask Nam if not.
4. Alt+F8 > `RegSetup` > Run.
5. Click **Pick register file** and choose the register on the shared drive.

## Daily use
1. Look at the panel in CreoRegister. Click **Refresh next free** to update it.
   Example: `M##1975` means "type M + your project + 1975".
2. In Creo: File > New > type the name (for example `M281975`) > in New File
   Options fill **DESCRIPTION** > OK.
3. In Excel: click **REGISTER ACTIVE CREO PART** or press **Ctrl+Shift+R**.
4. Check the confirm popup (name, description, sheet, project) > **Yes**.
5. "OK - registered" means the register is saved and closed. Done.

Keep the new part's window **active** in Creo when you click Register. The
tool registers whichever model is in the active window.

## What the tool will never do
- Write into a row that is already used, including rows that have a
  Project ID but no description.
- Write anything except Description (and Project ID on M/A).
- Change formats, formulas, locked cells or protection.
- Save the register if anything looks wrong. In that case it puts back the
  cells it wrote and closes the file without saving.
- Save or change anything in Creo. It only reads the model.

## Messages and what to do
| Message (short) | Meaning | What you do |
|---|---|---|
| Cannot connect to Creo | Creo is not running, two Creo sessions are open, or the API setup is missing | Start Creo, close any second session, then see Setup step 3 |
| Creo has no active model | No model window is active | Click into the new part's window in Creo |
| DESCRIPTION is blank / still START_PART | You didn't fill DESCRIPTION in New File Options | In Creo: File > Prepare > Model Properties > Parameters, then fill it |
| has no DESCRIPTION parameter (missing) | The part wasn't made from start_part | Add a DESCRIPTION parameter |
| must be exactly M + 2-digit project + 4-digit number | Name has a suffix or a typo | Rename the part in Creo |
| Project ID 'xx' is not listed | Project number typo, or a new project | Check the name; ask the register owner to add the project |
| number xxxx ... is already used | Someone already took that number | Click Refresh next free, rename the part in Creo |
| in use by someone else | A colleague has the register open | Wait until they close it, then click Register again |
| open with UNSAVED changes | You have the register open with edits | Save or close it first |
| no Part No formula | The prepared table ends before that number | Use the next free number, or ask the owner to extend the table |
| Read-back check failed ... (tool error) | The tool's safety check caught something unexpected. Nothing was saved | Report it to Nam with the full message |
| Unexpected error ... (tool error) | Bug in the tool | Report it with the full message |

**(missing)** in a message = your data needs fixing. **(tool error)** =
report it. Nothing is written when either one appears.

## Log
The **Log** sheet keeps every attempt: time, result, name, sheet, row,
project, description, reason and Windows user. Use it to find or undo an
entry by hand.

## Troubleshooting checklist
1. Module line 2 says `CreoRegister v1.0` (Alt+F11).
2. Debug > Compile VBAProject shows no error.
3. Exactly one Creo session is running.
4. The register path (cell B3) is correct, and the shared drive is connected.
5. Nobody else has the register open.
