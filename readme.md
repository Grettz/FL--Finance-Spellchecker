# Installation:
Requires [Python 3.8.6](https://www.python.org/downloads/) (May not work with other versions)

#### Install requirements
```
pip install requirements.txt
```

# Usage:
This is what your directory should look like before you run the program:
```
(optional) virtualenv/
(optional) remove_duplicates.py
dictionary.xlsx
text.xlsx
spellchecker.py
spell_check_text.py
apply_user_actions.py
```
##### dictionary.xlsx:
- Excel file containing all custom dictionary words on the first sheet and listed in the first column.
- Words can also be seperated in the same cell with pipes '|' e.g. 'stock | TSLA | dividends' etc.

##### text.xlsx:
- Excel file containing all the words/sentences to spellcheck.
- Text must be listed in the first worksheet and in the first column (A), or listed in any column with the column header 'Text'

##### Remove Duplicates
Run the 'remove_duplicates.py' script to delete all duplicate lines/sentences in the 'text.xlsx' file. Doing this before you run the main script can save significant time if there are many duplicates.

# Run the main script
```
spell_check_text.py
```
##### Optional Flags
`--no-suggestions` Disables auto-correct and word suggesions, outputting only spelling errors and performing the fastest by far.

`no--auto` Disables only auto-correct but fetches openpyxl suggestions. Performs ~20x slower than `--no-suggestions`.

No Flags - Fetches suggestions, auto-corrects words, and google searches words unable to correct. Performs ~20x slower than `--no-auto`.

`--debug` Creates a 'debug.xlsx' Excel file containing all words that have been checked.
##### e.g.
```
spell_check_text.py --no-auto
```

# Results Output
The program outputs a 'results.xlsx' file containing 2 worksheets. 

The first worksheet 'Non Corrected' contains all words that were not corrected by the program, and the second worksheet 'Corrected' contains all words corrected by the program, either through suggestions or from google.

Both worksheets contain the context for the spelling mistake with a hyperlink to a google seach, the line number of the sentence in the input 'text.xlsx' file, and a list of suggested words that the spelling mistake could be.

Both worksheets contain a 'User Action' column where you can modify what happens to the words/spelling mistakes before they are corrected.

#### User Actions
The user actions input behave slightly differently on the 'Non Corrected' and 'Corrected' worksheets. 

A blank User Action cell in the 'Non Corrected' worksheet will ignore the spelling mistake as there were no suggestions found that are likely to be the intended spelling; however a blank User Action in the 'Corrected' worksheet will apply the Suggested Word to the text output file when the 'apply_user_actions.py' script is ran.

The possible inputs for User Actions are as follows:
- 'a' or 'add': Adds the flagged/misspelt word to the dictionary, and does not change the word in the text file.

- 'd' or 'del': Deletes the word in the sentence in the text file.

- '1', '2', '3': Corrects the word using the given suggestion number.

- Any other text: Corrects the word to this string, allowing you to change the word to whatever word you need.
    
After you have made all the modifications you want, you need to run the 'apply_user_actions.py' script to apply all the word changes and output the 'text-output.xlsx' file, and add any words to the dictionary.

# Text Output
A 'text-output.xlsx' file is generated that is a copy of the input 'text.xlsx' file, but with all word corrections applied.
