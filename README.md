<!--lint disable no-literal-urls-->
<p align="center">
  <a href="https://Oso.com/">
    <img
      alt="Oso"
      src="http://clipart-library.com/images/ki8pMekir.png" width="192" height="120" 
    />
  </a>
</p>

“Selenium_Expedia” Is designed to show basic web test automation . 

Rest of readme coming soon!!!

# Table of contents

* [Support](#support)
* [Overview](#overview)
* The Process
  * [Imports](#imports)
  * [Web Scrapping](#web-scraping)
  * [PDF Conversion](#pdf-conversion)
  * [Data Cleaning](#data-cleaning)
  * [Data Export](#data-Export)

* [Current Project Team Members](#current-project-team-members) 

* [Collaborators](#collaborators)
  
## Support

Looking for help? Send an email for direct support &lt;hizstor@gmail.com&gt;

## Overview
SPC was built with PyCharm 2020.3.3 running Python 3.9.1.

SPC is designed as a first step to automating the process of compiling SAT data. It will scrape a specific website where individual state SAT data is maintained. The data is collected from the website, reformatted and combined into a single source that can easily be referenced and analyzed. This makes it simple to compare groups of SAT takers. Comparisons can be made between states, regions, demographic groups, previous years, etc. Having the data stored in this manner also makes visualization of the data simple and modular. 


## Imports
Libraries that must be imported to python environment for code to run successfully 

* **pandas**: Used to build DataFrames.
* **BeautifulSoup**: From bs4, used for web scraping.
* **requests**: Used for web scraping.
* **urljoin**: From urllib.parse, used for web scraping.   
* **tabula**: Used to extract data from PDFs.
* **glob, os**: Used to work with directories on a local machine. 


## Web Scraping

SPC will navigate to [CollegeBoard] website page that has a list of PDFs for each state and territory in the USA as well as a single PDF containing national data. Each file is identified by a .pdf href and downloaded to a local machine. The files are renamed at this time so they will be easier to reference in later steps. This will obtain the PDFs from the most recent reported year (currently 2020.) If earlier data sets are desired the code can easily be modified to collect that data from multiple years.
    
## PDF Conversion

The most challenging part of this project is converting information contained in a PDF to data easily accessed in python. The current code only extracts the data from a specific section of each PDF. This can be expanded as desired if information from other portions of the PDFs would be considered useful. SPC extracts this data by iterating through each PDF that was downloaded, identifying tables in the PDF and putting that data into a python list so it can be cleaned. 

## Data Cleaning

As is the case with most data obtained from outside sources, this data has a lot of erroneous or poorly organized information. SPC cleans the data so that it can more easily be utilized and understood. Erroneous data is removed and relevant data is reorganized in a way that can be used to create functional lists, dictionaries or DataFrames in Python. This part of the process also combines the data organized by state. 

## Data Export

Now that the data is clean it can be exported in the desired format. For this project it was exported as a csv file structured as a simple Python dictionary and as a csv file structured as a Pandas DataFrame. The code can easily be modified to export the data in other formats as desired.  

## Current project team members
* [OsoSuerte](https://github.com/OsoSuerte) -
**Ben Sadler** &lt;hizstor@gmail.com&gt; (he/him)


<!--lint disable prohibited-strings-->

### Collaborators

* [OsoSuerte](https://github.com/OsoSuerte) -
**Ben Sadler** &lt;hizstor@gmail.com&gt; (he/him)

<!--lint enable prohibited-strings-->

[CollegeBoard]: https://reports.collegeboard.org/sat-suite-program-results/
