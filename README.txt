================================================================================
                    PAG-IBIG HOUSING LOAN FORMS AUTO-FILL SYSTEM
               BLOOMFIELD TERRACES · BRGY. MAYBANCAL, MORONG, RIZAL
                      SUMMERFIELD RESOURCES CORPORATION
================================================================================

Welcome to the Pag-IBIG Forms Automation System. This tool automatically fills
and generates official Pag-IBIG Fund housing loan documents for buyers of
Bloomfield Terraces, pulling information directly from your shared BookData.xlsx.


--------------------------------------------------------------------------------
1. QUICK START FOR OFFICE USERS (WINDOWS PC)
--------------------------------------------------------------------------------

Option A: Run Immediately Without Installing
1. Double-click the file named:
      TRY-ON-PC.bat
2. The application will open on your screen. You will see the list of buyers
   loaded directly from BookData.xlsx.

Option B: Install as a Regular Windows Program
1. Double-click:
      _setup.bat
   (This sets up the required environment once.)
2. Double-click:
      BUILD-WINDOWS.bat
3. When finished, open the "dist" folder and run the .msi installer.


--------------------------------------------------------------------------------
2. USING THE SINGLE-FILE WEB TOOL (ANY BROWSER)
--------------------------------------------------------------------------------
If you prefer using a web browser (Google Chrome, Microsoft Edge, or Safari):
1. Navigate to the "web" folder.
2. Double-click:
      index.html
3. The modern glass-styled tool will open instantly.
4. Click "Load Excel (BookData.xlsx)" to select your file, or work directly
   with the interactive forms and preview the documents live.


--------------------------------------------------------------------------------
3. HOW TO GENERATE LOAN DOCUMENTS
--------------------------------------------------------------------------------
1. Find your buyer:
   - Use the Search box on the left to search by Name, Co-Borrower, Block &
     Lot, or TCT, then click the buyer's card.
2. Review or update details:
   - The details are grouped in four tabs (Parties & Property, Financials,
     Technical Description, Notary & Attorney-in-Fact).
   - Click "Save to Excel" (or press Ctrl+S) to update BookData.xlsx in place.
     "Unsaved changes" shows until you save.
   - Click "Preview" on any form to see exactly how it will print.
3. Choose the Stage:
   - Select either "PRE-APPROVAL" or "POST-APPROVAL".
   - The app will automatically determine whether to use the Affordable
     Housing Program (< ₱950,000) or End-User Program (₱950,000 and above).
4. Click "Save & Make Forms":
   - The buyer is saved to BookData.xlsx first, so the forms always match Excel.
   - The app lists any important details that are still blank and asks before
     continuing. Blank details stay blank on the forms (nothing is made up).
   - The PDFs will be created and saved into a dedicated folder named:
     "<BUYER NAME> - Pre-approval" or "<BUYER NAME> - Post-approval".
   - Quick "Open" buttons will appear next to each generated file.
5. Loanable Amount: always type the APPROVED amount. When an HLA is scanned, the
   amount the buyer asked for is shown as a note but is not filled in.


--------------------------------------------------------------------------------
4. WHERE FORMS ARE SAVED
--------------------------------------------------------------------------------
The first time the app opens on a computer, a setup screen asks for:
   - BookData.xlsx: pick the shared file (e.g. the OneDrive copy), or create
     a new empty one in a folder you choose (an existing file is never
     overwritten).
   - The folder with the official Pag-IBIG blank forms.
   - Where to save filled forms.
   - Developer details and, optionally, the Anthropic API key.
The app cannot be used until setup is finished. Everything can be changed
later in Settings.

If not changed, completed forms are saved in your computer's Documents folder:
   Documents \ Pag-IBIG Forms \ <BUYER NAME> - <Stage>

Click "Settings" at the top of the window to choose:
   - BookData.xlsx: the shared Excel file (for example the OneDrive copy).
   - Official blank forms folder: put the official Pag-IBIG blank PDFs here
     (e.g. HLF161_CTSFinancingAffordableHousingProgram_V06.pdf). If a form's PDF
     is missing, the app warns you and makes a DRAFT layout marked
     "DO NOT SUBMIT" instead.
   - Save filled forms in: where the buyer folders are created.

The Anthropic API key (for HLA scanning) is stored encrypted for your Windows
user on this computer. Use "Remove Saved API Key" in Settings to delete it.


--------------------------------------------------------------------------------
5. EXCEL FILE (BookData.xlsx) BEST PRACTICES & SAFETY
--------------------------------------------------------------------------------
• Keep Columns A to O in their original order.
• Always close BookData.xlsx before clicking "Save to Excel" so Windows
  and OneDrive allow the file to be updated. If it is open in Excel, the app
  will safely warn you: "Close the Excel file and try again".
• The app uses atomic saving: it will never delete rows, corrupt files, or lose
  data if the computer turns off or loses connection.


--------------------------------------------------------------------------------
6. OFFICE DISTRIBUTION & SECURITY GUIDELINES
--------------------------------------------------------------------------------
To protect company and buyer privacy:
1. Distribute Installers Only:
   Give office staff the generated installer (.msi or .apk), NEVER the source code.
2. Personal Installation:
   Install the software personally on authorized company computers only.
3. Shared Excel Access:
   Share BookData.xlsx on OneDrive/SharePoint with named, authorized team members
   only. Never use "Anyone with the link can edit".
4. Staff Offboarding:
   When an employee leaves the company, immediately remove their access from the
   shared OneDrive/SharePoint folder.
5. Backup & Recovery:
   If someone accidentally overwrites data in Excel, use OneDrive's "Version
   History" (right-click the file in OneDrive -> Version History) to restore
   previous copies in seconds.
6. Authorized Device Use:
   Have staff sign a standard company data-protection agreement for handling
   buyer confidential information.


--------------------------------------------------------------------------------
7. CONTACT & SUPPORT
--------------------------------------------------------------------------------
Developer: Summerfield Resources Corporation
Project: Bloomfield Terraces, Brgy. Maybancal, Morong, Rizal
Documentation Head: Jennifer B. Peduca
Technical Support: info@summerfield.ph
================================================================================
