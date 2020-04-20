# PDF to Word

Typora is a useful program to create documentation and instructions. It has built-in feature supporting latex to write technical documents. But Typora will not be able to generate headers and footers when converting to PDF. 

This project is to complement Typora by performs following:

* Convert PDF file to images
* Crop images to remove surrounding white space
* Insert images to Word document, either from existing template file or new file. 



## Setup

Install Python libraries

```bash
pip install -r requirements.txt
brew install poppler
```

