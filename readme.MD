
# Text File to Excel Converter

This project is a Python-based application that converts specific text file data into a structured Excel file. The application provides a graphical user interface (GUI) using PyQt5, allowing users to easily load a text file, process its content, and export the results to an Excel file.

## Features

- **File Input**: Users can input the text file via the GUI.
- **Data Processing**: The application processes the text file, calculating time differences and formatting the output.
- **Table Display**: Processed data is displayed in a table within the GUI.
- **Excel Export**: The processed data can be exported and saved to an Excel file.

## Installation

### 1. Clone the Repository

```bash
git clone https://github.com/raufur-simanto/text-to-excel-converter.git
cd text-to-excel-converter
```

### 2. Create and Activate a Virtual Environment

```bash
python3.10 -m venv venv
source venv/bin/activate  # On Windows use `venv\Scripts\activate`
```

### 3. Install the Required Dependencies

```bash
pip install -r requirements.txt
```

### 4. Create an executable file
```
pyinstaller --onefile main.py
```

## Usage

### 1. Run the Application

```bash
python main.py
```

### 2. Load Text File

- Enter the name of the text file or browse to select it.
- The file content will be processed, and the results will be displayed in the table within the GUI.

### 3. Export to Excel

- Click the "Export to Excel" button to save the processed data as an Excel file.

## Project Structure

- `main.py`: The main application file that runs the GUI.
- `requirements.txt`: List of Python dependencies.
- `README.md`: Project documentation.

## Dependencies

- Python 3.10+
- PyQt5
- pandas
- openpyxl

Install all dependencies using:
```bash
pip install -r requirements.txt
```

## Troubleshooting

### Common Issues

1. **Missing Python Libraries**:
   - If you encounter issues related to missing Python libraries or shared objects, ensure that your Python environment is properly set up with the necessary development packages.

2. **PyInstaller Issues**:
   - If using `PyInstaller` to package the application, ensure that Python is built with `--enable-shared` if you encounter `PythonLibraryNotFoundError`.

### Solutions

- Install Python development packages:
  ```bash
  sudo apt-get install python3.10-dev
  ```

- Enable shared library if `PythonLibraryNotFoundError`
  ```bash
  cd Python-3.10.x
  ./configure --enable-shared
  make
  sudo make install
  ```

- Update `LD_LIBRARY_PATH` if the shared library is not found:
  ```bash
  export LD_LIBRARY_PATH=/usr/local/lib:$LD_LIBRARY_PATH
  ```

## Contributing

Contributions are welcome! Please fork this repository, make your changes, and submit a pull request.

---

This `README.md` file provides a complete overview of your project, including installation instructions, usage, and troubleshooting tips. Make sure to adjust any specific details, such as the repository URL, to fit your actual project setup.