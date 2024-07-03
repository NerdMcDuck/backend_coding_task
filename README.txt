This is a backend coding test I did a couple years ago. It has been refactored and updated.

List of Stopwords

Source: https://gist.github.com/sebleier/554280

Location: https://gist.githubusercontent.com/rg089/35e00abf8941d72d419224cfd5b5925d/raw/12d899b70156fd0041fa9778d657330b024b959c/stopwords.txt

Steps:

	> Create a virtual env (venv, conda, etc.) and activate it
	> pip install requirement.txt

Notes:

The "test docs" folder should be in the same directory as the python file. 
My program only takes ".txt" files, so to check more files than the included,
we just need to add the .txt file to the "test docs" folder.

To Run:
	python  .\coding_task.py

After Run:
	This will generate an Excel file 'InterestingWords.xlsx' with the table in
	the following format:

	Word (Total Occurrence) | Document | Sentences containing the word
	
	I chose the top 5 words from each document and 3 sentences that used the
	 corresponding word from the document it was found in.
	
	I've included the output file as well.
