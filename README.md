The code parses the PDF file declaration and saves the data to a DataFrame.

The interface is very simple. First, you select a PDF file and click OK. Below is an example of the working window

![image](https://github.com/user-attachments/assets/38f8c1b5-1c88-45be-b940-099522b793e8)

The DataFrame is saved in Excel format and looks like this:

![image](https://github.com/user-attachments/assets/ca93b6b7-431b-48df-83d9-460ca4b3254e)

Additionally, the program logs information to a .log file. It works quite efficiently, taking approximately 2 minutes to parse the PDF file with 460 pages.

The example of .log file:

2024-07-27 20:08:55,614 INFO File D:/41.pdf is open
2024-07-27 20:08:55,614 INFO Start from main
2024-07-27 20:08:55,706 INFO Thread started
2024-07-27 20:08:55,707 INFO Start executing
2024-07-27 20:08:55,989 INFO The line 1 added to df

2024-07-27 20:10:04,081 INFO The line 688 added to df
2024-07-27 20:10:04,093 INFO End executing, decode to excel
2024-07-27 20:10:05,488 INFO Prog ended
