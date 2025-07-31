
## BenfordsLaw_Tester.py 
This script models Benford’s Law by generating and comparing authentic versus manipulated financial data. It outputs frequency distributions to Excel for forensic analysis, helping identify statistical anomalies suggestive of fraud."



Installation:
```
python pip install -r requirements_benfords.txt
```

Usage:\
analyze numbers
```
python BenfordsLaw_Tester.py -c -I benfordsLaw_tester.xlsx
```


help menu
```
python BenfordsLaw_Tester.py -h
```

Example:

    BenfordsLaw_Tester.py -b
	
    BenfordsLaw_Tester.py -b -I benfordsLaw_tester.xlsx
	
Note:
this currently reads column A in the first sheet.
	
	
![sample output](images/BenfordsLaw_Enron.png)	
	


## Manipulated Data
The Manipulated Data section simulates financial values that are artificially rounded 
to the nearest thousand, creating unnaturally uniform distributions. These values do not 
follow the logarithmic pattern predicted by Benford’s Law, which typically governs organic 
datasets. By comparing the first-digit frequency of this manipulated data to Benford's 
expected distribution, the model helps demonstrate how fabricated or tampered numbers 
diverge from statistical norms.

## Authentic Data
The ## Authentic Data section contains numerical values that are naturally occurring 
and statistically organic. These figures are either real-world samples or generated 
using randomization techniques that closely mimic legitimate datasets. They tend to 
follow Benford’s Law, which predicts the frequency of first digits in many 
naturally formed datasets.

## Chi-Square Test
The Chi-Square Test is a statistical method used to evaluate whether observed frequencies 
in a dataset differ significantly from expected frequencies under a certain hypothesis. 
It's commonly applied to categorical data—like survey responses, classifications, or 
groupings—to test for independence or goodness-of-fit.
The bigger the number, the less likely it is that the differences occurred by chance, 
and the more likely it is that something meaningful is happening in the data.

Chi-Square Test ➤ Stat: 1114.61



