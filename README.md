This is a standalone application that takes a Microsoft Word .docx file input and removes all blacklines before creating a separate copy with "(No Blacklines) " prepended to the front. (The original file is preserved). Features include:

1. Drag & drop
2. Support for multiple files at once
3. Lightning fast speed, even for larger documents
4. option between .docx and .pdf outputs
5. color customization for which font color you want 1. exclusively deleted 2. deleted for strikethroughs and cleaned for underlines and 3. exclusively cleaned (cleaned meaning color and underlining removed)

The .exe should be a runnable standalone program. When you run the program, you will get a pop-up saying "Windows Protected Your PC." You will have to press on the "More Info" button, then "Run Anyway." If the .exe runs too slowly, or Windows flags the programs as a false positive, you will need to install Python:

1. Open the Microsoft Store app by pressing the search bar in the task bar -> type in "Microsoft Store" -> open the app
2. In the app, search for "Python" and click on the most recent version (Python 3.13 as of this readme)
3. Press "Get"
4. Once the application downloads, press Windows Key + R -> type "CMD" (without the quotes) -> press enter (alternatively, press the search bar in the task bar and look for 5. "CMD" -> press on "Command Prompt")
5. When the command line opens, type (without quotations) "Python -m pip install PyQt6 lxml docx2pdf" (this should install all the necessary dependencies) and press enter
6. run the program by double clicking on BlacklineRemover.py

Please contact me if issues persist or if you find any bugs/unwanted removals/suggestions.
