INSTRUCTIONS:
1. Download the Python application on website
2. Create a master study spreadsheet (.csv) that contains 3 columns:
    a. Column A: the catogories and years: I.e. Phase I -> Healthy Adults -> (year)
    b. Column B: The study name, protocol name, and description. Can include treatment names
    c. Column C: The study name and description; remove protocol name and treatment name; but could be the same as Column B if user doesn't care.
2. Double click on the file called RUN.bat
3. In “One-Click (CSV No-Year Fix), browse and select 2 files:
	3a. CV (.docx): the person’s original CV.
	3b. CSV (.csv) file: A spreadsheet that contains 3 columns: year (column A), red label study (column B), and non-red label study (column C)
Alternative (if no CSV file): Use the “Three Files (+ Splitter)” tab that takes in 3 inputs:
	a. CV (.docx): the person’s original CV.
	b. Unred master study list (.txt): the plain-text file containing all study records without red labels.
        1. Red labels = protocol name and treatment name
	c. Red-labeled master study list (.docx): the Word document containing the same studies but with colored / formatted text (red protocol labels).
3.5. (Optional): Check the box to split the CV into two pdf files: Abbrv and non-Abbrv.
4. Click “Run”
5. When complete, the updated CV will appear in the Output folder.

How it works: the program will automatically:
1. Extract studies from the CV (“Research Experience” section).
2. Sort and match them using the unred master study list (.txt).
3. Merge in new studies and preserve red labels from the red master (.docx).
4. Inject the finalized list back into the original CV.
5. When complete, the updated CV will appear in the folder. 
- If the checkbox to split is checked, three .docx will appear in that folder: updated unsplit CV, the Abbrv CV (first half) and UnAbbrv CV (second half)


Note & Rules:
1. The script reads studies below “Research Experience” and above this sentence in the CV:
“By signing this form, I confirm that the information provided is accurate and reflects my current qualifications.”
* It removes the old list in between and injects the new studies in that same location.

2. The non-red master study's must follow this rule: YEAR is followed by a TAB -> the PROTOCOL followed by the colon ":". 
For example: 
YEAR<TAB>PROTOCOL: Description of the study
"2025	ABBVIE: A Phase 1 Multiple Ascending Dose, Drug-Drug Interaction, and Asian Pharmacokinetic Study of XXXX  "
- Look at the NO_RED_STUDYLIST in Editable folder as a sample.
3. The red master study must follow the same rule as above. Year -> Protocol -> ":" -> Description.
* The red text (protocol name or study treatment) will be preserved when merged into the final CV.
* The final updated CV will follow the red text document format. If you see inconsistent spacing, fix it in the red document so it auto-updates when inserting.

4. If the CV doesn't have any studies under research experience, put the year of the CV last submitted (i.e. 2022) right under Research Experience and it will auto=populate studies after 2022 (so studies from 2023 to current). 
