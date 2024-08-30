# Console application DepartmentModulesDescription

This is app that get root folder and handle department folders with university data which contains .docx files with university subjects and their desctiption. To every speciality there is .xlsx file with working hours and count of exams. In result we get .docx files by every speciality about subjects that will be studied with name, count exams and description


*Tags* : DocumentFormat.OpenXml.Wordprocessing, OfficeOpenXml, sqlite-net-pcl, GTranslatorAPI, MethodTimer.Fody, Microservices

Step 1
-
WordFileReader read .docx files and try to find table with data about speacialities that read name and description of subject. In succesess case data write into local database.

Step 2
-
ExcelExecuter write into every .xlsx file of speaciality description to studied subjects.

Step 3
-
ExcelExecuter read total data from .xlsx and give data to WordFileWriter and it generates output files according to template formatting.
