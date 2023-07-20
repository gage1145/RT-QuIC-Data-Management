 _______________    ________\n
|  _\__   __/   \ _ \__   _ \ \n
| |__) | | |  _  | | | | | \_| \n
|  __ <| |=| [_] | |_| | |  _ \n
| |  \ | | |     |\__^_| |_/ | \n
|_|  |_|_|  \__\_\  \_______/ \n

README

This program is designed to handle data exported from MARS into a .xlsx file report.

RT_xlsx_organize3.R will read the second sheet in the .xlsx file which contains the real-time data. It will then do two things:
1) Organize replicate columns so that samples are side-by-side.
2) Create separate sheets for each read type (e.g. Raw, Normalized, Derivative, etc.)

Each sheet is then exported and saved into the original file, so no extra files are created, and no data is overwritten.

RT_xlsx_organize3.R will automatically download and install any libraries you will need. If they are already installed, it will do nothing.

There are 3 steps of user input:
1) Enter the working directory. This should include forward slashes and not back slashes (e.g. "/" not "\"; C:/Users/username/Documents)
2) Enter the file name WITHOUT the file extension. The program will ask again if a file is not found. (e.g. "rt-quic_2023" not "rt-quic_2023.xlsx")
3) Enter the number of rows of metadata present on sheet 2 of the Excel file. This may be different depending on the amount of data you want exported by MARS.

The original file will be opened automatically, and you can check that your data has been exported properly.

Troubleshooting

1) If you notice that your time values are not starting at "0", or your column identifiers are incorrect, try reducing the number of rows of metadata you want removed. The program may have removed too much.
