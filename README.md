# Excel Column Extractor (XLSColumnFilter)

**Excel Column Extractor** is a web-based tool that allows users to upload an Excel file and extract specific columns from it. This tool uses [SheetJS](https://github.com/SheetJS/sheetjs) to handle Excel file parsing and creation. It offers a user-friendly interface with smooth animations and an elegant design.

## Features

- Upload `.xlsx` or `.xls` files.
- Extract specific columns based on column headers.
- Download the filtered data as a new Excel file.
- Beautiful and smooth UI animations using CSS.


<!-- https://github.com/user-attachments/assets/bfb17bcb-06e0-4b78-b675-38c922d50f94 -->


## Demo

![](https://github.com/epsit03/XLSColumnFilter/blob/main/assets/ScreenRecording2024-09-14193401-ezgif.com-speed.gif)

## How It Works

The user selects an Excel file, and the script extracts only the desired columns based on the provided column headers. The filtered data is then downloaded as a new Excel file.

## Getting Started

To run this tool, simply clone the repository and open `index.html` in a web browser.

### Prerequisites

You need a web browser with internet access to load the [SheetJS](https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js) library.

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/epsit03/ExcelColumnExtractor.git

## ToDo (Future Changes To Be Done...)
- Enabling the user to dynamically input the columns he/she wants to fetch alongwith its data
- Also, enabling the user to specify on which header are the columns located. (For e.g. If the column headers are in the first row, user will input header position = 1)
- Make the UI more smooth and elegant.
```As for now, this web-based tool is specific for an organisation's purpose but my purpose to make it open-source is to make the tool viable and useful for everyone across the global web.```

## Feel free to open a PR if you have any prospective ideas ;-)
