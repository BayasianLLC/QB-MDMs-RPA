# QB MDM Automation Project

## Overview
This project automates the process of monitoring, downloading, and transforming PSEG MDM (Master Data Management) files from SharePoint to CSV format for QuickBase integration.

## Features
- Automated SharePoint monitoring for new MDM files
- XLSB to CSV conversion
- Data transformation and cleanup
- Continuous monitoring with configurable intervals
- Detailed logging and error handling

## Prerequisites
- Python 3.12 or higher
- Anaconda or Miniconda
- SharePoint access credentials
- Required Python packages:
  - pandas
  - pyxlsb
  - Office365-REST-Python-Client

## Installation

1. Clone the repository:
```bash
git clone https://github.com/BayasianLLC/QB-MDMs-RPA.git
cd QB-MDMS-RPA
```

2. Create and activate a Conda environment:
```bash
conda create -n qb_mdm python=3.12
conda activate qb_mdm
```

3. Install required packages:
```bash
conda install -c conda-forge Office365-REST-Python-Client
conda install pandas pyxlsb
```

## Configuration

1. Update SharePoint credentials in the script:
```python
sharepoint_url = "your_sharepoint_url"
username = "your_email"
password = "your_password"
```

2. Configure file paths:
```python
folder_path = "your_sharepoint_folder_path"
output_folder = "your_local_output_path"
```

## Usage

Run the script:
```bash
python src/mdm_processor.py
```

The script will:
1. Monitor SharePoint for new MDM files
2. Download and process new files when detected
3. Convert files to CSV format
4. Save transformed files to the specified output location

## File Structure
```
PSEG-MDM-RPA/
├── src/
│   └── mdm_processor.py
├── config/
│   └── settings.py
├── logs/
├── README.md
└── requirements.txt
```

## Troubleshooting

Common issues and solutions:

1. SSL Certificate Errors:
```bash
conda config --set ssl_verify true
conda install ca-certificates
```

2. SharePoint Connection Issues:
- Verify credentials
- Check SharePoint URL format
- Ensure proper folder paths

## Contributing
1. Fork the repository
2. Create your feature branch
3. Commit your changes
4. Push to the branch
5. Create a new Pull Request

## License

MIT License

Copyright (c) 2024 BayasianLLC

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

## Author
BayasianLLC

## Support
For support, please open an issue in the GitHub repository.