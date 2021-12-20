# Example

This is the excel file used to develop this program. The original excel file can be found here: https://www.nrel.gov/grid/solar-resource/clear-sky.html 

Original license can be found in this folder under "NREL_Data_Disclamer", or here: https://www.nrel.gov/disclaimer.html


The output.py file was generated with this command line:
```
python main.py bird_08_16_2012.xlsx 0 --alternating-def A8:A29 -z B2:Z2,B3:Z3 | black - > output.py
```

To see the resulting code in use, take a look at test_output.py