# CV Generator

Generate customized CVs from templates with placeholder replacement.

## Setup

### Using uv (recommended)

```bash
# Create virtual environment
uv venv

# Activate it
source .venv/bin/activate

# Install dependencies
uv pip install -r requirements.txt
```

### Using pip

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Usage

### Command-line Options

*   `--template <path>`: Path to the CV template file (required).
*   `--output <path>`: Path to save the filled CV (required).
*   `--data <path>`: Path to JSON file with replacement data.
*   `--pdf`: If present, also generates a PDF file from the DOCX.
*   `--role <string>`: Specify the role for organizing output files.
*   `--company <string>`: Specify the company name for organizing output files.

### Using JSON data file
```bash
.venv/bin/python3 cv_generator.py --template example_template.docx --output my_cv.docx --data example_data.json
```

### Using command-line arguments
```bash
.venv/bin/python3 cv_generator.py --template example_template.docx --output my_cv.docx -D NAME="Jane Smith" EMAIL="jane@example.com"
```

### Mix JSON and command-line (command-line takes precedence)
```bash
.venv/bin/python3 cv_generator.py --template example_template.docx --output my_cv.docx --data example_data.json -D NAME="Jane Smith"
```

### Generating PDF with role and company specific output
```bash
.venv/bin/python3 cv_generator.py --template example_template.docx --output my_cv.docx --data example_data.json --pdf --role "QA Engineer" --company "Nationale Nederlanden Group"
```

## Template Format

Create your CV template as a `.docx` file with placeholders in the format: `{{PLACEHOLDER}}`

Examples:
- `{{NAME}}`
- `{{EMAIL}}`
- `{{PHONE}}`
- `{{EXPERIENCE}}`

## Data Format (JSON)

```json
{
  "NAME": "John Doe",
  "EMAIL": "john@example.com",
  "PHONE": "+1 (555) 123-4567",
  "LOCATION": "San Francisco, CA"
}
```

## Files

- `cv_generator.py` - Main script
- `example_template.docx` - Sample DOCX template
- `example_data.json` - Sample data file
- `requirements.txt` - Python dependencies
