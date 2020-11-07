# OFFERRUM Parser
The parser receives information about the offers of the CPA network OFFERRUM and creates an article about these offers in DOCX and HTML formats.

Installation
---
The parser itself doesn't require installation, but it needs some dependencies to work. You have to install:

1. Python 3.6 or above;
2. Firefox browser (preferably the latest version);
3. [geckodriver](https://github.com/mozilla/geckodriver/releases) — its directory must be in the PATH or the *geckodriver* executable file must be in the same directory as the parser;
4. Some Python packages required for the parser to work. You can see the list of these packages in the `requirements.txt` file. You can install them using the `pip install -r requirements.txt` command in Terminal or Command Prompt (running Terminal or Command Prompt in the same directory as the `requirements.txt` file).

Usage
---
After the installation of all the dependencies, you can use the parser.

First, write the links to the necessary offers to the `links.txt` file. The file should contain links like these:
```
https://my.offerrum.com/offers/21173
https://my.offerrum.com/offers/21002
https://my.offerrum.com/offers/21166
```

Then you can run the parser. It will ask your email and password to enter the OFFERRUM website, and also the article's filename (where to save the result).
