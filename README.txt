Future updates planned: 1) Add graphical user interface (GUI)
                        2) Improve sensitivity to unusual searches such as "mOtion".
			            3) Add extra data regarding keywords, such as line number and times appearing in documents.

RatiFinder Version 0.2
"Automated Word File Directory Search"
Authors: Christian Pearson and Christian Pearson (Nov 2 2018)

Intended use: Automate identification of key sentences in Toastmasters Meeting Minutes Word files.
              Target strings present in the Minutes files are typically in the format of:

                MOTION: X motions that Y. Z in favour. Passed/Not passed.

Process: RatiFinder requests user input to define the keyword (item_type) to search for,
then opens and parses Word (.docx) files in the current directory for the immediately associated
information from each file, starting from the beginning of each instance of the inputted keyword
to the end of the line it occurs on.

All instances of a motion and its corresponding text are converted into a list and outputted as .txt file.
Program is useful for searching large numbers of Word Files for search words to quickly locate relevent information.


Known bugs:

- Unusual searches such as mixed case terms ("aCTion") may cause program to skip over the data that should have
been found.