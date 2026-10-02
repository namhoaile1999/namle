# CreoRegister v1.0 - Test Protocol (run once, in this order)

Status of v1.0: **BUILT, NOT YET COMPILED OR RUN.** I could not compile VBA or
connect to Creo from the build environment. Nothing below has been run yet.
Every step is a real test.

Use a **COPY** of the register for steps 1-8. Never use the real register
until step 9.

## Preparation (about 5 minutes)
1. Copy `Part_number_register.xlsx` from the shared drive to a local test
   folder, e.g. `C:\CreoRegTest\Part_number_register.xlsx`.
2. Open a new blank Excel workbook. **File > Save As > Excel Macro-Enabled
   Workbook (.xlsm)** and save it as `C:\CreoRegTest\CreoRegister.xlsm`.
   Save it BEFORE importing the code, so the tool can save its log.
3. Alt+F11 > File > Import File > choose `CreoRegister.bas`.
4. Check that line 2 of the module says **`CreoRegister v1.0`**. If it doesn't,
   you imported an old or renamed download.
5. Tools > References > tick **Creo VB API Type Library**. Tick exactly ONE
   Creo library.
6. **Debug > Compile VBAProject.** Expect no message. If you get an error,
   STOP and send me the exact text and the highlighted line.
7. Close the VBA editor. Alt+F8 > `RegSetup` > Run. Expect a "setup done" popup.

## Step 1 - Panel + Creo connection proof
1. Click **Pick register file** and choose the COPY.
2. Expected panel:

| Sheet | Next free number | Type in Creo as | Register row | Prepared rows left | Last used description |
|---|---|---|---|---|---|
| Manufactured parts (M) | 1975 | M##1975 | 2042 | 2014 | 30 MM valve module P2 |
| Assemblies (A) | 0561 | A##0561 | 579 | 176 | SP Stepped Knurled scew Beta ASM |
| Bought parts (B) | B002678 | B002678 | 1618 | 420 | Small Diameter Ball Bearing 8x14x4 |

   PASS = all 18 values match. Any difference: send a screenshot.
   The numbers shift if your copy is newer than the file you sent me; the
   pattern is what matters.

## Step 2 - Happy path, Manufactured (this also proves the two new Creo API calls)
1. In Creo: File > New > Part > name `M281975`, template start_part,
   DESCRIPTION = `TEST PART M` > OK.
2. In Excel press **Ctrl+Shift+R** (this tests the shortcut too).
3. Expect a confirm popup showing Name M281975, Description TEST PART M,
   Sheet Manufactured parts, Project 28. Click **Yes**.
4. Expect "OK - registered ... row 2042".
5. Open the COPY by hand and check row 2042: B = **M281975**, C = TEST PART M,
   D = 28. Then close it WITHOUT saving.
   - If instead you see "Could not get the active model" or a compile error
     on `CurrentModel` / `InstanceName`, the unverified API names are the cause.
     Send me the exact text.

## Step 3 - Happy path, Assembly and Bought
1. New assembly `A280561`, DESCRIPTION `TEST ASM` > Register > Yes.
   Expect row 579.
2. New part `B002678`, DESCRIPTION `TEST BOUGHT` > Register > Yes.
   Expect row 1618. B has no project line.

## Step 4 - Click twice
With `B002678` still active, press Ctrl+Shift+R again > Yes.
Expect **"Already registered ... nothing written"**.

## Step 5 - Broken cases (each one must say NOT registered, with a clear reason, and change nothing)
| # | What to do in Creo | Expected message mentions |
|---|---|---|
| a | New part `M280053`, DESCRIPTION `X` | number 0053 already used, Project ID 03, Description (blank) |
| b | New part `M281976`, leave DESCRIPTION = `START_PART` | still the template default |
| c | New part `M991976`, DESCRIPTION `X` | Project ID '99' is not listed on the Index sheet |
| d | New part `E213167`, DESCRIPTION `X` | Elec parts are not handled |
| e | New part `M281976_V2`, DESCRIPTION `X` | must be exactly M + 2-digit project + 4-digit number |
| f | Assembly `A280740`, DESCRIPTION `X` (row has a number but no Part No formula) | no Part No formula ... cleared again |

After a-f, open the COPY and confirm none of those rows changed. Close
without saving.

## Step 6 - Register busy (simulates a colleague holding the file)
1. In Windows Explorer, right-click the COPY > Properties > tick
   **Read-only** > OK.
2. New part `M281976`, DESCRIPTION `TEST BUSY` > Register > Yes.
   Expect "in use by someone else (or you have no write permission) ...
   Nothing was written".
3. Untick Read-only again.

## Step 7 - Log
Open the tool's **Log** sheet. Expect one row per attempt from steps 2-6,
with Time, Result (OK / ALREADY / FAIL / CANCELLED), name, sheet, row and
reason.

## Step 8 - Quality check on the copy
Open the COPY by hand:
- Rows 2042 (M), 579 (A) and 1618 (B) hold the test entries. Every other row
  is unchanged.
- Sheet protection is still on: try typing in a locked Part No cell, and it
  must refuse.
- Delete the 3 test entries by hand if you want to reuse the copy.

## Step 9 - Real register
Only after steps 1-8 pass: click Pick register file > choose the REAL
register on the shared drive. Delete the test parts in Creo.

## What to send me
For each step: PASS, or FAIL plus the exact popup text or a screenshot.
If step 6 behaves differently (for example the save goes through), also
answer: **is the register on a mapped network drive (S:\ or \\server\...)
or in SharePoint / Teams / OneDrive?**
