
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

    BenfordsLaw_Tester.py -b -I benfordsLaw_tester.xlsx -c b
		
Note:
this currently reads column A in the first sheet use -c to specify a differnt column.
	
	
![sample output](images/BenfordsLaw_Enron.png)	
	
## Benford's Law (Blue line)
Benford's Law, also known as the First-Digit Law, is a statistical principle that predicts 
the frequency distribution of leading digits in naturally occurring datasets. 

According to this law:

• 	The number 1 appears as the leading digit about 30% of the time.

• 	Higher digits occur less frequently, with 9 appearing as the first digit less than 5% of the time.

This counterintuitive pattern holds true across many datasets—such as financial records, 
population numbers, and scientific measurements—especially when the data spans several orders of magnitude.

Benford's Law is widely used in fraud detection, particularly in accounting and forensic analysis. 
Deviations from the expected distribution can signal manipulation or anomalies in the data.
The blue line gives a baseline of what Benford's law should look like, as a baseline.

## Chi-Square Test
The Chi-Square Test is a statistical method used to evaluate whether observed frequencies 
in a dataset differ significantly from expected frequencies under a certain hypothesis. 
It's commonly applied to categorical data—like survey responses, classifications, or 
groupings—to test for independence or goodness-of-fit.
The bigger the number, the less likely it is that the differences occurred by chance, 
and the more likely it is that something meaningful is happening in the data.

Chi-Square Test ➤ Stat: 1114.61

## Mean Absolute Deviation (MAD)
MAD measures how far the observed digit proportions deviate from Benford’s expected distribution — on average. 
It’s a simple yet powerful way to quantify irregularities: the higher the MAD, the more the data strays 
from what Benford’s Law predicts. In forensic analysis, a low MAD suggests natural, unmanipulated data, 
while a high MAD can signal potential anomalies worth investigating.



