# Grade Entry Tool

A small CLI tool designed to simplify entering student grades directly into Excel.

The program allows fast keyboard-based grade entry by searching students by name or ID, supports optional class-based filtering using seating plan PDFs, and writes grades directly into the selected Excel file.

---

### Installation

```bash
git clone https://github.com/oguz-duman/grading
```

```bash
pip install -r requirements.txt
```

```bash
python main.py
```

---

### Usage

1. **Excel File**

    The program works directly with an Excel file (.xlsx) containing the student list.
    The file must include at least:

    * student ID column
    * student name column
    * grade column

    The program will ask for the starting cell of each column (e.g. A2, B2, C2).
    <br>
    <br>

2. **Class-based Filtering** (Optional)

    To simplify grade entry, students can be filtered by the classroom in which they took the exam.

    Seating plan PDFs may be placed in `class_filters/`.


    The program extracts numeric values from each PDF and interprets them as student IDs.

    Each PDF file must contain the seating plan of exactly one class.

    File names are not used programmatically, but they are displayed when selecting the active filter.
    <br>
    <br>    


3. **Program Flow**

    When the program starts, the following steps are performed:
    <br>
    <br>

    a. **Select Excel file**

    A file dialog is opened to select the Excel file containing the student list.
    <br>
    <br>

    b. **Column configuration**

    The program asks for the starting cells of the relevant columns.
    
    For example:
    * A2  → student ID
    * B2  → student name
    * C2  → grade
    <br>
    <br>

    c. **Existing grades**

    If the Excel file already contains grades, they are preserved.

    Students with an existing grade are automatically skipped.

    This also makes it possible to stop grading midway and continue later without losing progress.
    <br>
    <br>

    d. **Grade range**

    The program asks for minimum and maximum allowed grade values to prevent invalid input.

    Both integer and decimal values are supported.
    <br>
    <br>

    e. **Selecting students**

    A searchable student list is displayed.

    Students can be located by typing part of `student ID` or `student name`

    Grades can then be entered directly.

    Once a student's grade is entered, the student is removed from the search results to make further searches easier.

    The most recently entered grades are displayed at the top of the screen, allowing quick verification.
    <br>
    <br>

    f. **Canceling a selection**

    If a student is selected accidentally, entering one of the following values instead of a grade cancels the selection:

    * `q`
    * `exit`
    * `cancel`
    * `break`
    <br>
    <br>

    g. **"..." element**

    The `...` element in the students list exists only to prevent accidental key presses.

    It has no functional role.
    <br>
    <br>

    h. **Menu options at the bottom of the student list**

    Most of the control options are located below the student list and may not be immediately noticeable.

    They can be reached quickly using the arrow keys:

    `Change Filter`, `Finish Grading`.
    <br>
    <br>

    j. **Change Filter Option**

    If filter PDFs are present, the active filter can be changed during grading.

    This allows focusing on one classroom at a time.
    <br>
    <br>

    k. **Finish Grading Option**
    
    Selecting Finish ends the grading session.

    All entered grades are written directly into the Excel file.

    Since students with an existing grade in the Excel file are automatically skipped, it is possible to continue grading later without losing progress.

---

### Output


Grades are written directly into the selected Excel file.

Existing grades remain unchanged.


The original row order is preserved.

Students without a grade remain unchanged.

---

### Limitations

Grades cannot be modified once they are entered.

If an incorrect grade is entered, the recommended workflow is:

1. Note down the student ID and the correct grade
2. Continue grading
3. Correct the grade manually in Excel afterwards

Direct editing inside the program is intentionally not supported in order to keep the input process fast and simple.