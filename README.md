# Word Generator

This project generates Word documents (`.docx`) using a template and data from an Excel or CSV file.

## Project Structure
- `generate_docs.py`: Main script to generate documents.
- `data.xlsx` / `data.csv`: Input data file (Excel or CSV).
- `template.docx`: Word template used for document generation.
- `generated/`: Output folder for generated Word documents (customizable).
- `requirements.txt`: Python dependencies.

## Installation

1. **Clone the repository**
   ```powershell
   git clone <repository-url>
   cd word-generator
   ```

2. **Install Python (if not already installed)**
   Download and install Python from [python.org](https://www.python.org/downloads/).

3. **Install dependencies**
   ```powershell
   pip install -r requirements.txt
   ```

## Usage

1. Prepare your input data in an Excel (`.xlsx`) or CSV (`.csv`) file and your template in `template.docx`.
2. Run the script:
   ```powershell
   python generate_docs.py <data_file> <template_file> --output-dir <output_directory>
   ```
   - Example:
     ```powershell
     python generate_docs.py data.xlsx template.docx --output-dir generated
     ```
   - The `--output-dir` argument is optional. Default is `generated_docs`.
3. Generated documents will appear in the specified output folder.

### Template Placeholders
- In your Word template, use placeholders in the format `<column_name>`.
- Each placeholder will be replaced with the corresponding value from the input data for each row.

## Requirements
- Python 3.7+
- See `requirements.txt` for required packages.