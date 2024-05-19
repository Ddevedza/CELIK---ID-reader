# Celik-ID Reader

## Overview
Celik-ID Reader is a Python application designed to interface with electronic ID cards using Serbia's official Celik API. The application reads data from an ID card inserted into a card reader and performs several functions, including data extraction, document generation, and data storage.

## Features
- **Data Extraction**: Reads data from electronic ID cards using the Celik API.
- **Document Generation**: Fills a predefined Word template with extracted data for printing or archival.
- **Data Storage**: Saves extracted ID data into an Excel spreadsheet to keep records that can be easily accessed and analyzed.
- **User Interface**: Provides a basic GUI to initiate the ID card reading process and display operation status.

## Requirements
- Python 3.x
- Windows operating system (due to dependency on Celik API and COM components)
- Required Python libraries: `ctypes`, `tkinter`, `openpyxl`, `docxtpl`, `win32com.client`

## Usage

To start the application, run the following command from the project directory:

```bash
python Main2.py
```
The GUI will launch, and you can begin by inserting an ID card into the reader and clicking the **"Istampaj podatke"** button.

## How It Works

### Initialization
- On launch, the application initializes by setting up the necessary directories and checking for the required dependencies.

### Reading ID Card
- When an ID card is inserted, the application uses the Celik API to read data from the card.

### Data Processing
- Extracted data is used to fill a Word document template.
- Relevant data points are also stored in an Excel file for record-keeping.

### Error Handling
- The application includes basic error handling capabilities to manage and troubleshoot potential issues during the read process.

## Contributing

Contributions to the Celik-ID Reader are welcome! Please read `CONTRIBUTING.md` for details on our code of conduct, and the process for submitting pull requests to us.

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE) file for details.

## Acknowledgments

- Thanks to the Serbian Ministry of Interior for providing the Celik API.

