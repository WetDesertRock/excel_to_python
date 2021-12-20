# Excel to Python

Excel to Python is a proof of concept/learning program written to transform algorithms written in excel documents to python code.

## Process overview

1. User provides excel worksheet, and the ranges of cells that define the algorithm and constants/variables ued in the algorithm
2. The program reads through the given ranges and attempts to assign variable names to the provided cells.
3. The program goes through each formula cell and transforms it to python code:
  A. The openpyxl tokenizes the excel formula
  B. The resulting tokens are run through a shunting yard algorithm which transforms the tokens into python ast
  C. The python ast then gets fed through an ast.NodeTransformer which will transform the excel functions into python equivalents.
4. The resulting python ast's are combined with the remaining cell variables and fed into a function which will combine everything into a python function

## FAQ

### What can and can't this do?

This can convert the excel file found here: https://www.nrel.gov/grid/solar-resource/clear-sky.html

Any excel features not used by this algorithm will not have been implemented.

### Should I use this code?

Probably not.

### What needs to happen next in this project?

 - Refactor the code involved in steps 1-4.
 - Add in automatic code testing to ensure the resulting code generates output matching what excel had calculated previously
 - Expand capabilities.

### The resulting code is ugly!

Oh no!

Feed it through an automatic code formatter such as black:
```s
python main.py bird_08_16_2012.xlsx 0 --alternating-def A8:A29 -z B2:Z2,B3:Z3 | black - > output.py
```