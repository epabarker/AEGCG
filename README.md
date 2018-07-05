# AEGCG
## Automated Excel Grade Calculator Generator. Generates an excel spreadsheet to calculate your grades based upon input of modules and module details. 

### Plenty of grade calculators exist already, why another?

The existing grade calculators I have seen online are numerous, and many are good. However, they have many problems. First and foremost, they do not allow "storage" of your grades. Many people like to keep track of their grades over time, and for this reason it was important to me to create a calculator that was exportable/downloadable, in a format friendly to everyone. .xlsx(Excel) is ideal for this, as 99% of people have Excel, and have used it before. 

Secondly, many calculators online do not allow for separation of marks by module. All modules are weighted, and not necessarily uniformly. To list all assessments and their weights is not sufficient, further weighting is applied by module. It is also useful to have both grade averages and overall grades for each module, as this is something people often look at. 

Every person's course is unique, and this grade calculator generator hopes to account for that. 

The whole point of a grade calculator is the utility to the user, which is what I have designed this for. It includes module averages, overall module percentage, overall module classification, semester average; percentage; classification, and yearly average; percentage; classification. 

This is currently a work in progress, and more features are coming. 

### Information you will need about your course to use this program:

- Number of credits in year
- Name of each module
- Number of credits in each module
- Number of pieces of coursework per module
- Type of coursework
- Weighting of each piece of coursework

### How to use:

Clone the github, or navigate to the directory where the repository has been downloaded, then run the excelGradeMaker.py in the terminal by doing the following command:
```
$ python excelGradeMaker.py
```
You will then be greeted with a series of questions regarding the structure of your degree. After all questions have been concluded, an excel spreadsheet will be generated in the same directory as the python file. 

### NOTE:
xlsxwriter is not able to AutoFit, as this can only be done at runtime. Make sure to AutoFit Columns when you open your file. 


To be added:

  - ~~OOP~~
  - ~~Good coding style~~
  - GUI
  - Web application
  - Input validation
  - Multiple academic years
  - Semesters?
  - ~~Degree classification, average per semester, average for year~~
  - Download file prompt
  - Download/view counter
