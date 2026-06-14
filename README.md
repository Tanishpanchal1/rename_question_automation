Approach Overview
Instead of sorting the alphanumeric suffixes alphabetically (which would randomize the question sequence relative to the actual exam layout), we use the Excel file as the source of truth:

- Extract Mapping: Load the spreadsheet and map each QBG Question id (hash) to its corresponding Display Order* (e.g., 1.0, 2.0, etc.).
- Locate Files: For each question ID, search the image directory for its matching question image (QUES_ENG_<id>.<ext>) and solution image (SOLU_ENG_<id>.<ext>).
- Copy & Rename: Clean/recreate the renamed_assets folder and copy the paired files over, renaming them to Q{order}.{ext} and S{order}.{ext} respectively (e.g. Q1.png, S1.png).
